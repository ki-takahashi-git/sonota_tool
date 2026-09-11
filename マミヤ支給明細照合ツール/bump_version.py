# -*- coding: utf-8 -*-
"""
bump_version.py  -  VERSIONファイルのビルド番号(パッチ番号)を1つ上げる

EXE化.bat がビルドのたびに実行し、"MAJOR.MINOR.PATCH" の PATCH だけを
自動で増やす(他のツールと同じ方式)。MAJOR/MINORは大きな変更を
したときに手動でVERSIONファイルを書き換えること。
"""

import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_PATH = os.path.join(_HERE, "VERSION")


def main():
    with open(_PATH, encoding="utf-8") as f:
        ver = f.read().strip()

    parts = (ver.split(".") + ["0", "0", "0"])[:3]
    major, minor, patch = parts
    try:
        patch = str(int(patch) + 1)
    except ValueError:
        patch = "1"
    new_ver = "%s.%s.%s" % (major, minor, patch)

    with open(_PATH, "w", encoding="utf-8") as f:
        f.write(new_ver + "\n")

    print("Version bumped: %s -> %s" % (ver, new_ver))


if __name__ == "__main__":
    main()
