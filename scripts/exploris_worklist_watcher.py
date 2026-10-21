import argparse, csv, hashlib, json, logging, re, shutil, subprocess, tempfile, time
from pathlib import Path, PureWindowsPath
import openpyxl

HEADERS = ['Sample Type','File Name','Sample ID','Path','Instrument Method','Process Method','Calibration File','Position','Inj Vol','Level','Sample Wt','Sample Vol','ISTD Amt','Dil Factor','L1 Study','L2 Client','L3 Laboratory','L4 Company','L5 Phone','Comment','Sample Name','']
PAT = re.compile(r'^(EXP\d{2}-\d{5}-\d{3})(?:$|[ -].*)')
METHOD = r'D:\Production_Data\PremierLC\12min_8CM_PremLC_TE_targeted_uLPickup'

def run(a, *args):
    r = subprocess.run([a.rclone,*args,'--config',a.config,'--retries','2','--timeout','60s','--contimeout','15s'],capture_output=True,text=True,encoding='utf-8',timeout=240)
    if r.returncode: raise RuntimeError(r.stderr)
    return r.stdout

def generate(book, name, a):
    wb=openpyxl.load_workbook(book,read_only=True,data_only=True)
    samples=[]
    try:
        rows=wb['Sample List'].iter_rows(values_only=True)
        for number,row in enumerate(rows,1):
            h=[str(v).strip() if v is not None else '' for v in row]
            if 'Well' in h and 'Peptide name' in h: break
        else: raise ValueError('Missing sample headers')
        for rownum,row in enumerate(rows,number+1):
            d={k:v for k,v in zip(h,row) if k}
            peptide=str(d.get('Peptide name') or '').strip()
            if not peptide: continue
            well=str(d.get('Well') or '').strip()
            if not well and peptide.lower()=='you fill in': continue
            if not re.fullmatch(r'[A-H](?:[1-9]|1[0-2])',well): raise ValueError(f'Invalid position at row {rownum}: {well!r}')
            d.update({'Well':well,'Peptide name':peptide,'Source row':rownum,'Source sheet':'Sample List'})
            samples.append(d)
    finally: wb.close()
    if not samples: raise ValueError('No samples')
    if len({d['Well'] for d in samples})!=len(samples): raise ValueError('Duplicate positions')
    identity=PAT.fullmatch(name).group(1)
    target=str(PureWindowsPath(a.local_root)/name)
    sequence=[]; metadata=[]; seen=set()
    for d in samples:
        well=d['Well']; short=re.sub(r'[^A-Za-z0-9_.-]+','_',d['Peptide name']).strip('_.')
        filename=f'{identity}_{well[0]}{int(well[1:]):02d}_{short}_I01'.replace('.', 'p')
        if filename in seen or len(filename)>150 or len(target+'\\'+filename+'.raw')>240: raise ValueError('Duplicate or overlong filename')
        seen.add(filename)
        entry={k:'' for k in HEADERS}
        entry.update({'Sample Type':'Unknown','File Name':filename,'Sample ID':d.get('Sample ID') or short,'Path':target,'Instrument Method':a.method,'Position':well,'Inj Vol':2,'Sample Wt':0,'Sample Vol':0,'ISTD Amt':0,'Dil Factor':1,'L1 Study':identity,'Comment':'; '.join(f'{k}: {d[k]}' for k in ('Notes','Sequence') if d.get(k)),'Sample Name':d['Peptide name']})
        sequence.append(entry)
        metadata.append(dict(d,**{'File Name':filename,'RAW filename':filename+'.raw','Source workbook':book.name}))

    # Filename policy: replace periods with p
    # Injection policy: G1/G5 blanks 4uL; samples 2uL
    blank_rows = []
    blank_metadata = []

    for position in ('G1', 'G5'):
        blank_name = f'{identity}_Blank_{position}_I01'

        if blank_name in seen:
            raise ValueError(f'Duplicate filename: {blank_name}')
        seen.add(blank_name)

        blank = dict(sequence[0])
        blank.update({
            'Sample Type': 'Unknown',
            'File Name': blank_name,
            'Sample ID': f'Blank_{position}',
            'Position': position,
            'Inj Vol': 4,
            'Comment': 'Solvent blank',
            'Sample Name': f'Blank {position}',
        })
        blank_rows.append(blank)

        blank_metadata.append({
            'Well': position,
            'Peptide name': f'Blank {position}',
            'Sample ID': f'Blank_{position}',
            'Notes': 'Solvent blank; added by generator',
            'Source sheet': 'Generator',
            'Source row': '',
            'File Name': blank_name,
            'RAW filename': blank_name + '.raw',
            'Source workbook': '',
            'Injection volume (uL)': 4,
        })

    for item in metadata:
        item['Injection volume (uL)'] = 2

    sequence = blank_rows + sequence
    metadata = blank_metadata + metadata
    digest=hashlib.sha256(book.read_bytes()+json.dumps([name,a.method,target,'v5_periods_p_blanks4_samples2']).encode()).hexdigest()
    destination=Path(a.workspace)/'generated'/name/digest[:16]
    if destination.exists():
        logging.info('Already generated: %s',destination); return
    destination.parent.mkdir(parents=True,exist_ok=True)
    temp=Path(tempfile.mkdtemp(dir=destination.parent,prefix='building-'))
    try:
        with (temp/'worklist_DRAFT.csv').open('w',encoding='utf-8-sig',newline='') as f:
            f.write('Bracket Type=4,\r\n')
            w=csv.DictWriter(f,fieldnames=HEADERS);w.writeheader();w.writerows(sequence)
        with (temp/'sample_metadata.csv').open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=list(dict.fromkeys(k for d in metadata for k in d)));w.writeheader();w.writerows(metadata)
        shutil.copy2(book,temp/'source_workbook.xlsx')
        (temp/'manifest.json').write_text(json.dumps({'sha256':digest,'samples':len(samples),'injections':len(sequence),'status':'DRAFT - review before acquisition'},indent=2))
        temp.rename(destination)
    finally:
        if temp.exists(): shutil.rmtree(temp)
    logging.info('GENERATED %d injections including 2 blanks: %s',len(sequence),destination)

def poll(a):
    folders=json.loads(run(a,'lsjson',a.remote,'--dirs-only'))
    names=[f['Name'].casefold() for f in folders]
    for f in folders:
        name=f['Name']
        if not PAT.fullmatch(name): continue
        try:
            if re.search(r'[<>:"/\\|?*\x00-\x1f]',name) or name.endswith((' ','.')): raise ValueError('Unsafe Windows folder name')
            if names.count(name.casefold())!=1: raise ValueError('Duplicate Drive folder name')
            local=Path(a.local_root)/name
            if local.is_symlink() or (local.exists() and not local.is_dir()) or (local.exists() and getattr(local,'is_junction',lambda:False)()): raise ValueError('Unsafe local destination')
            local.mkdir(parents=True,exist_ok=True)
            remote=a.remote.rstrip('/')+'/'+name
            files=json.loads(run(a,'lsjson',remote,'--files-only'))
            books = [
                x for x in files
                if x['Name'].lower().endswith('.xlsx')
                and not x['Name'].startswith('~$')
                and not x['Name'].lower().startswith('sample_sheet_snapshot_')
            ]
            if len(books)!=1: raise ValueError('Ready folder must contain exactly one .xlsx')
            item=books[0]
            with tempfile.TemporaryDirectory(dir=Path(a.workspace)/'cache') as td:
                book=Path(td)/Path(item['Name']).name
                run(a,'copyto',remote+'/'+item['Name'],str(book))
                after=json.loads(run(a,'lsjson',remote,'--files-only'))
                latest=[x for x in after if x['Name']==item['Name']]
                if len(latest)!=1 or any(item.get(k)!=latest[0].get(k) for k in ('ID','Size','ModTime')): raise ValueError('Workbook changed during download')
                generate(book,name,a)
        except Exception: logging.exception('Folder failed: %s',name)

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--once',action='store_true')
    p.add_argument('--remote',default='gdrive:Exploris480/Internal_Experiments')
    p.add_argument('--local-root',default=r'D:\Production Data\Internal_Experiments')
    p.add_argument('--workspace',default=r'D:\Scripts\ExplorisWorklists')
    p.add_argument('--rclone',default=r'D:\Scripts\rclone\rclone.exe')
    p.add_argument('--config',default=r'D:\Scripts\rclone\rclone.conf')
    p.add_argument('--method',default=METHOD)
    a=p.parse_args()
    Path(a.workspace).mkdir(parents=True,exist_ok=True)
    (Path(a.workspace)/'cache').mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s',handlers=[logging.StreamHandler(),logging.FileHandler(Path(a.workspace)/'watcher.log',encoding='utf-8')])
    lock=(Path(a.workspace)/'watcher.lock').open('a+b')
    import msvcrt
    lock.seek(0);lock.write(b'0');lock.flush();lock.seek(0)
    try: msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
    except OSError: raise SystemExit('Another watcher is running')
    for location in (a.rclone,a.config,a.local_root):
        if not Path(location).exists(): raise SystemExit('Missing path: '+location)
    while True:
        try: poll(a)
        except Exception: logging.exception('Poll failed')
        if a.once: break
        time.sleep(60)


# Publish generated documents alongside RAW and mzML files.
_original_generate = generate

def generate(book, name, a):
    _original_generate(book, name, a)

    target = Path(a.local_root) / name
    raw_path = str(PureWindowsPath(a.local_root) / name)
    digest = hashlib.sha256(
        book.read_bytes()
        + json.dumps([name, a.method, raw_path, 'v5_periods_p_blanks4_samples2']).encode()
    ).hexdigest()

    source = Path(a.workspace) / 'generated' / name / digest[:16]
    revision = digest[:16]

    documents = {
        'worklist_DRAFT.csv': f'worklist_DRAFT_{revision}.csv',
        'sample_metadata.csv': f'sample_metadata_{revision}.csv',
        'source_workbook.xlsx': f'sample_sheet_snapshot_{revision}.xlsx',
        'manifest.json': f'worklist_manifest_{revision}.json',
    }

    target.mkdir(parents=True, exist_ok=True)

    for original, published in documents.items():
        src = source / original
        dst = target / published
        if not src.is_file():
            raise FileNotFoundError(f'Missing generated document: {src}')
        if dst.exists():
            logging.info('Preserving existing document: %s', dst)
            continue

        with dst.open('xb') as output:
            with src.open('rb') as input_file:
                shutil.copyfileobj(input_file, output)

        logging.info('SAVED beside data: %s', dst)


if __name__=='__main__': main()
