import argparse
import hashlib
import io
import pathlib
import re
import subprocess
import sys
import tokenize
import xml.etree.ElementTree as ET
import zipfile


ROOT = pathlib.Path(__file__).resolve().parent.parent
ALLOWED = {
    '.gitattributes', '.gitignore', 'LICENSE', 'NOTICE', 'README.md', 'SECURITY.md',
    'docs/PRIVACY.md', 'docs/images/main-window.png', 'docs/images/photo-preview.png',
    'src/Core.cs', 'src/Security.cs', 'src/UI.cs', 'src/ViewLayout.cs', 'src/MergeWorkflow.cs',
    'src/WordLayout.cs', 'src/Program.cs', 'src/LetterMerger.csproj', 'src/app.manifest',
    'tests/Tests.cs', 'tests/TestRunner.cs', 'tests/LetterMerger.Tests.csproj',
    'scripts/Build.ps1', 'scripts/Test.ps1', 'scripts/AuditRepository.py',
    'Runtime/LetterMerger.exe', 'Runtime/START_HERE.md',
    'Runtime/Data/README.md', 'Runtime/Photos/README.md', 'Runtime/Templates/README.md',
    'Runtime/Templates/A7_Card_Template.docx', 'Runtime/Output/README.md'
}
REVIEWED_BINARY = 'adfa65f6e1b147f81b486f3efe293423e0ce4bef11d5fd7a5ff1f685213b368c'
REVIEWED_TEMPLATE = '81690f6657562a3d61f36d5ba050fdf135fa5a15096160d1f8c4e7f259a0809a'
REVIEWED_LICENSE = '3972dc9744f6499f0f9b2dbf76696f2ae7ad8af9b23dde66d6af86c9dfb36986'
REVIEWED_IMAGES = {
    'docs/images/main-window.png': 'cbbac4a4bf8fcda1e71030ae279e1ebb13a9d1d4365c7b955a4b59bc8635e40e',
    'docs/images/photo-preview.png': '7b3409882c3bd311cd81827f5af5ce8aa26c813a7b083e34ac25721596feafd8'
}
PATTERNS = {
    'private key': r'-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----',
    'GitHub credential': r'\b(?:gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{30,})\b',
    'AWS access identifier': r'\b(?:AKIA|ASIA)[A-Z0-9]{16}\b',
    'Slack credential': r'\bxox[baprs]-[A-Za-z0-9-]{20,}\b',
    'Google API credential': r'\bAIza[A-Za-z0-9_-]{35}\b',
    'OpenAI credential': r'\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{32,}\b',
    'JWT credential': r'\beyJ[A-Za-z0-9_-]{15,}\.[A-Za-z0-9_-]{15,}\.[A-Za-z0-9_-]{15,}\b',
    'social security number': r'\b[0-9]{3}-[0-9]{2}-[0-9]{4}\b',
    'private workstation path': r'(?:[A-Za-z]:\\' + 'Users' + r'\\|/' + 'home' + r'/|/' + 'workspace' + r'/|/' + 'Users' + r'/)',
    'credential-bearing URL': r'https?://[^\s/]+:[^\s/]+@'
}


def fail(path, reason):
    raise ValueError(f'{path}: {reason}')


def git(*arguments):
    result = subprocess.run(['git', *arguments], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode:
        raise ValueError('Git could not read the staged file inventory. Stage the reviewed files first.')
    return result.stdout


def has_code_comments(source, powershell=False):
    i = 0
    while i < len(source):
        c = source[i]
        if (powershell and (c == '#' or source.startswith('<#', i))) or (not powershell and (source.startswith('//', i) or source.startswith('/*', i))):
            return True
        if c not in "\"'":
            i += 1
            continue
        quote = c
        verbatim = not powershell and i > 0 and source[i - 1] == '@'
        i += 1
        while i < len(source):
            if (powershell and quote == '"' and source[i] == '`') or (not powershell and not verbatim and source[i] == '\\'):
                i += 2
            elif source[i] == quote:
                if i + 1 < len(source) and source[i + 1] == quote and (verbatim or (powershell and quote == "'")):
                    i += 2
                else:
                    i += 1
                    break
            else:
                i += 1
    return False


def audit_text(path, content):
    try:
        text = content.decode('utf-8-sig')
    except UnicodeDecodeError:
        fail(path, 'expected UTF-8 text')
    if chr(8212) in text:
        fail(path, 'literal em dash found')
    for label, pattern in PATTERNS.items():
        if re.search(pattern, text):
            fail(path, label + ' pattern found; matching content was not printed')
    if re.search(r'\bCo-authored-by\s*:', text, re.I):
        fail(path, 'coauthor trailer found')
    if path.endswith('.cs') and has_code_comments(text):
        fail(path, 'C# comment found')
    if path.endswith('.ps1') and has_code_comments(text, powershell=True):
        fail(path, 'PowerShell comment found')
    if path.endswith('.py'):
        if any(token.type == tokenize.COMMENT for token in tokenize.generate_tokens(io.StringIO(text).readline)):
            fail(path, 'Python comment found')
    if path.endswith(('.csproj', '.manifest')):
        ET.fromstring(content)
        if '<!--' in text:
            fail(path, 'XML code comment found')


def audit_template(path, content):
    if hashlib.sha256(content).hexdigest() != REVIEWED_TEMPLATE:
        fail(path, 'blank template differs from the reviewed template')
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        names = archive.namelist()
        if len(names) != 14 or len(set(names)) != len(names):
            fail(path, 'unexpected template part inventory')
        for part in names:
            data = archive.read(part)
            if part.endswith(('.xml', '.rels')):
                audit_text(path + ':' + part, data)
                ET.fromstring(data)
        core = ET.fromstring(archive.read('docProps/core.xml'))
        if len(core) != 0:
            fail(path, 'author or history metadata remains')
        for part in names:
            if part.endswith('.rels'):
                for relation in ET.fromstring(archive.read(part)):
                    if relation.attrib.get('TargetMode', '').lower() == 'external':
                        fail(path, 'external template relationship')


def audit_files(files):
    if set(files) != ALLOWED:
        unknown = sorted(set(files) - ALLOWED)
        missing = sorted(ALLOWED - set(files))
        raise ValueError(f'File inventory differs from the reviewed allowlist. Unknown: {unknown}. Missing: {missing}.')
    for path in sorted(files):
        content = files[path]
        if len(content) > 2 * 1024 * 1024:
            fail(path, 'unexpected file size')
        if path.endswith('.exe'):
            if hashlib.sha256(content).hexdigest() != REVIEWED_BINARY or not content.startswith(b'MZ'):
                fail(path, 'executable differs from the reviewed build; review a rebuilt executable before updating the approved hash')
        elif path.endswith('.png'):
            if hashlib.sha256(content).hexdigest() != REVIEWED_IMAGES.get(path) or not content.startswith(b'\x89PNG\r\n\x1a\n'):
                fail(path, 'screenshot differs from the reviewed image')
        elif path.endswith('.docx'):
            audit_template(path, content)
        else:
            audit_text(path, content)
        if path == 'LICENSE' and hashlib.sha256(content).hexdigest() != REVIEWED_LICENSE:
            fail(path, 'GPL version 3 license text differs from the reviewed full license')
        if path == 'NOTICE' and b'SPDX-License-Identifier: GPL-3.0-only' not in content.splitlines():
            fail(path, 'the GPL-3.0-only project license notice is missing')


def main():
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--working-copy', action='store_true')
    mode.add_argument('--tracked', action='store_true')
    mode.add_argument('--committed', action='store_true')
    args = parser.parse_args()
    if args.tracked or args.committed:
        entries = git('ls-tree', '-r', '-z', '--full-tree', 'HEAD') if args.committed else git('ls-files', '--stage', '-z')
        files = {}
        for entry in entries.split(b'\0'):
            if not entry:
                continue
            metadata, raw_path = entry.split(b'\t', 1)
            path = raw_path.decode('utf-8')
            if metadata.split(b' ', 1)[0] not in {b'100644', b'100755'}:
                fail(path, 'only regular files may be published')
            files[path] = git('show', ('HEAD:' if args.committed else ':') + path)
        if args.committed:
            audit_text('commit message', git('log', '-1', '--format=%B'))
    else:
        files = {}
        for path in ROOT.rglob('*'):
            relative = path.relative_to(ROOT).as_posix()
            parts = pathlib.PurePosixPath(relative).parts
            if parts[0] in {'.git', 'TestResults'} or relative.startswith(('src/bin/', 'src/obj/', 'tests/bin/', 'tests/obj/', 'scripts/__pycache__/')):
                continue
            if path.is_symlink():
                fail(relative, 'filesystem link is unsupported')
            if path.is_file():
                if relative not in ALLOWED:
                    fail(relative, 'file is outside the reviewed allowlist; keep runtime data and generated results outside the repository')
                files[relative] = path.read_bytes()
    audit_files(files)
    print(f'PASS: {len(files)} reviewed files; no forbidden files, known secret patterns, inherited template author metadata, code comments, or literal em dashes detected.')
    print('This audit supplements manual review. It cannot identify every secret or personal fact, assess institutional authorization, or certify FERPA compliance.')


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('AUDIT FAILED: ' + str(error), file=sys.stderr)
        sys.exit(1)
