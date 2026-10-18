"""Read-only, isolated desktop rendering. Registered is not the same as usable."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import re
from report_core import digest, write_json

PROGIDS={'word':'Word.Application','wps':'kwps.Application'}

def normalize_headings(value):
    """PowerShell can wrap function-returned arrays inside an extra array."""
    if isinstance(value,dict):return [value]
    if value is None:return []
    if isinstance(value,list):return [h for item in value for h in normalize_headings(item)]
    raise ValueError('Invalid editor heading metadata')

def registration(backend):
    result={'backend':backend,'progid':PROGIDS[backend],'registered':False,'automation':'not_probed','version':None}
    if sys.platform!='win32':return {**result,'reason':'Windows COM unavailable'}
    import winreg
    for view,name in [(winreg.KEY_WOW64_64KEY,'64'),(winreg.KEY_WOW64_32KEY,'32')]:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,PROGIDS[backend]+r'\CLSID',0,winreg.KEY_READ|view) as k:clsid=winreg.QueryValue(k,None)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,'CLSID\\'+clsid+r'\LocalServer32',0,winreg.KEY_READ|view) as k:server=winreg.QueryValue(k,None)
            result.update(registered=True,registry_view=name,server=server)
            # Match the server's bitness, especially 32-bit WPS on 64-bit Windows.
            shell=Path(os.environ.get('WINDIR',r'C:\Windows'))/('SysWOW64' if name=='32' else 'System32')/'WindowsPowerShell/v1.0/powershell.exe'
            if shell.is_file():result['powershell']=str(shell)
            match=re.match(r'"([^"]+\.exe)"|(.+?\.exe)(?:\s|$)',server,re.I)
            if match and shell.is_file():
                exe=next(x for x in match.groups() if x)
                command="(Get-Item -LiteralPath '"+exe.replace("'","''")+"').VersionInfo.FileVersion"
                try:
                    p=subprocess.run([str(shell),'-NoProfile','-NonInteractive','-Command',command],capture_output=True,timeout=15,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
                    if p.returncode==0:result['file_version']=p.stdout.decode('utf-8',errors='replace').strip()
                except (OSError,subprocess.TimeoutExpired):pass
            return result
        except OSError:continue
    return {**result,'reason':'No complete COM server registration; installation or automation component may be missing'}

PS=r'''param([string]$Source,[string]$Pdf,[string]$Metadata,[string]$ProgId,[string]$ProcessName,[string]$RoundTrip)
$ErrorActionPreference='Stop'
$app=$null; $doc=$null; $owned=$false
function HeadingData($d) {
  $items=@()
  foreach($p in $d.Paragraphs) {
    $style=[string]$p.Style.NameLocal
    if($style -match '^SZUHeading[1-5]$') {
      $items+=@{text=$p.Range.Text.Trim([char[]]@(13,7,32,10));label=$p.Range.ListFormat.ListString;level=[int]$p.OutlineLevel;page=$p.Range.Information(3)}
    }
  }
  return ,$items
}
try {
  $prior=@(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
  $app=New-Object -ComObject $ProgId
  $beforeIds=@($prior | ForEach-Object {$_.Id})
  $created=@(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object {$_.Id -notin $beforeIds})
  if($app.Documents.Count -ne 0) {throw 'ISOLATION_BLOCKED: Automation reference has existing documents.'}
  # Empty Word has no document window. Open only our read-only temporary copy,
  # without changing application options, then identify that document's window.
  $doc=$app.Documents.Open($Source,$false,$true,$false)
  if([string]$doc.FullName -ne $Source) {throw 'ISOLATION_BLOCKED: Unexpected document identity.'}
  $instanceProcess=0
  try {
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class SzuWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$windowProcess=0
    $handle=[IntPtr]::new([long]$doc.Windows.Item(1).Hwnd)
    [SzuWindow]::GetWindowThreadProcessId($handle,[ref]$windowProcess)|Out-Null
    $instanceProcess=[int]$windowProcess
  } catch { Write-Output ('Window identity unavailable: '+$_.Exception.Message) }
  $created=@(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Where-Object {$_.Id -notin $beforeIds})
  $identified=($instanceProcess -gt 0 -and $instanceProcess -in @($created | ForEach-Object {$_.Id}))
  if(-not $identified -and $prior.Count -eq 0 -and $created.Count -eq 1) {$identified=$true}
  if(-not $identified -or $app.Documents.Count -ne 1) {
    @{version=[string]$app.Version;isolation='blocked';document_count=$app.Documents.Count} | ConvertTo-Json | Set-Content -LiteralPath $Metadata -Encoding UTF8
    @{before_process_ids=$beforeIds;new_process_ids=@($created|ForEach-Object {$_.Id});instance_process_id=$instanceProcess;document_count=$app.Documents.Count;version=[string]$app.Version} | ConvertTo-Json -Compress | Write-Output
    throw 'ISOLATION_BLOCKED: Cannot prove exclusive ownership; no existing document/settings were changed, and the instance will not be closed automatically.'
  }
  $owned=$true
  $app.Visible=$false; $app.DisplayAlerts=0
  $version=[string]$app.Version
  $doc.Repaginate()
  $headings=HeadingData $doc
  $doc.ExportAsFixedFormat($Pdf,17)
  $result=@{version=$version;pages=$doc.ComputeStatistics(2);headings=@($headings);isolation='exclusive_empty_instance';roundtrip=$null}
  if($RoundTrip) {
    try {$doc.SaveAs2($RoundTrip,12)} catch {$doc.SaveAs($RoundTrip,12)}
    $mode=0;$doc.Close([ref]$mode);$doc=$null
    $doc=$app.Documents.Open($RoundTrip,$false,$true,$false)
    $doc.Repaginate()
    $result.roundtrip=@{headings=@(HeadingData $doc);pages=$doc.ComputeStatistics(2);saved_copy=$RoundTrip}
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Metadata -Encoding UTF8
} finally {
  $mode=0
  # Close only the exact temporary document opened above, even on failed isolation.
  try {if($null -ne $doc -and ([string]$doc.FullName -eq $Source -or [string]$doc.FullName -eq $RoundTrip)){$doc.Close([ref]$mode)}}
  finally {if($owned -and $null -ne $app -and $app.Documents.Count -eq 0){$app.Quit([ref]$mode)}}
  if($null -ne $app){[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
}
'''

def render(path, output, backend='auto', soffice=None, dpi=200, roundtrip=False):
    if not 96<=dpi<=400:raise ValueError('--dpi must be between 96 and 400.')
    path=Path(path).resolve();out=Path(output).resolve()
    if out.exists() and any(out.iterdir()):raise ValueError('Render output directory must be empty.')
    out.mkdir(parents=True,exist_ok=True)
    backend='wps' if backend=='auto' and sys.platform=='win32' else ('libreoffice' if backend=='auto' else backend)
    before=digest(path);copy=out/'render-input.docx';shutil.copy2(path,copy);pdf=out/'report.pdf'
    result={'backend':backend,'source_sha256':before,'headings':[],'dpi':dpi,'version':None,'visual_review':'pending'}
    try:
        if backend in PROGIDS:
            reg=registration(backend)
            if not reg['registered'] or not reg.get('powershell'):raise RuntimeError(backend+': '+reg.get('reason','PowerShell host missing'))
            script=out/'editor-render.ps1';script.write_text(PS,encoding='utf-8-sig')
            args=[reg['powershell'],'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(script),'-Source',str(copy),'-Pdf',str(pdf),'-Metadata',str(out/'editor.json'),'-ProgId',PROGIDS[backend],'-ProcessName','WINWORD' if backend=='word' else 'wps']
            if roundtrip:args+=['-RoundTrip',str(out/'roundtrip.docx')]
            proc=subprocess.run(args,capture_output=True,timeout=180,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            (out/'renderer.log').write_bytes(proc.stdout+b'\n'+proc.stderr)
            if proc.returncode or not pdf.is_file():raise RuntimeError(backend+' export blocked/failed; inspect renderer.log. No fallback counts as validation of this editor.')
            result.update(json.loads((out/'editor.json').read_text(encoding='utf-8-sig')))
            # PowerShell may serialize empty/single collections inconsistently.
            result['headings']=normalize_headings(result.get('headings'))
            if result.get('roundtrip'):
                hs=normalize_headings(result['roundtrip']['headings'])
                result['roundtrip']['headings']=hs
                key=lambda a:[(h['text'],h['label'],h['level']) for h in a]
                result['roundtrip']['numbering_stable']=key(hs)==key(result['headings'])
                result['roundtrip']['sha256']=digest(out/'roundtrip.docx')
        elif backend=='libreoffice':
            if not soffice or not Path(soffice).is_absolute() or not Path(soffice).is_file():raise RuntimeError('Pass absolute bundled --soffice path; no system fallback.')
            args=[str(soffice),'-env:UserInstallation='+(out/'lo-profile').as_uri(),'--headless','--convert-to','pdf','--outdir',str(out),str(copy)]
            proc=subprocess.run(args,capture_output=True,timeout=180)
            (out/'renderer.log').write_bytes(proc.stdout+b'\n'+proc.stderr)
            made=copy.with_suffix('.pdf')
            if proc.returncode or not made.is_file():raise RuntimeError('Bundled LibreOffice export failed.')
            made.replace(pdf)
        else:raise ValueError('Unknown rendering backend.')
        if digest(path)!=before or digest(copy)!=before:raise RuntimeError('Source or read-only copy was modified during rendering.')
        from pypdf import PdfReader
        from pdf2image import convert_from_path
        from report_visual import inspect_pages
        reader=PdfReader(pdf);images=convert_from_path(str(pdf),dpi=dpi)
        texts=[p.extract_text() or '' for p in reader.pages]
        for i,im in enumerate(images,1):im.save(out/f'page-{i}.png')
        from report_tools import render_risks
        risks=render_risks(images,texts,result['headings'])+inspect_pages(images,out,dpi)
        result.update(pages=len(images),page_text=texts,risks=risks,pdf_sha256=digest(pdf),
                      page_image_sha256={f'page-{i}.png':digest(out/f'page-{i}.png') for i in range(1,len(images)+1)})
        write_json(out/'render.json',result);return result
    except Exception as exc:
        write_json(out/'failure.json',{'backend':backend,'status':'blocked','source_sha256':before,'error':str(exc)})
        raise

def doctor(probe=False, output=None):
    result={'editors':[registration(e) for e in ('wps','word')],'default_primary':'wps','default_compatibility':'word'}
    if probe:
        if output is None:raise ValueError('--probe requires --output-dir for diagnostic artifacts.')
        from docx import Document
        out=Path(output);out.mkdir(parents=True,exist_ok=True)
        fixture=out/'probe.docx'
        if fixture.exists():raise ValueError('Probe output directory already used.')
        d=Document();d.add_paragraph('SZU renderer capability probe');d.save(fixture)
        for e in result['editors']:
            try:
                r=render(fixture,out/e['backend'],e['backend'])
                e.update(automation='export_verified',version=r['version'])
            except Exception as exc:
                e.update(automation='blocked',reason=str(exc))
                metadata=out/e['backend']/'editor.json'
                if metadata.exists():e['version']=json.loads(metadata.read_text(encoding='utf-8-sig')).get('version')
    return result
