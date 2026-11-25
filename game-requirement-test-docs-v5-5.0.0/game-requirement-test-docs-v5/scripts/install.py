"""Copy a skill to a skills directory, refusing all overwrite."""
import argparse, os, shutil
from pathlib import Path

def install(parent):
    source=Path(__file__).resolve().parents[1]
    parent=Path(parent).expanduser().resolve(); target=parent/source.name
    if target.exists():raise SystemExit(f'同名技能已存在，未覆盖：{target}')
    if parent==source or source in parent.parents:raise SystemExit('目标不能位于源技能内部')
    parent.mkdir(parents=True,exist_ok=True)
    shutil.copytree(source,target,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
    if not (target/'SKILL.md').is_file():raise SystemExit('复制后缺少SKILL.md')
    print(f'已复制技能：{target}')

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    default=Path(os.environ['CODEX_HOME'])/'skills' if os.environ.get('CODEX_HOME') else Path.home()/'.codex/skills'
    p.add_argument('--skills-dir',default=str(default),help='目标skills父目录')
    install(p.parse_args().skills_dir)
