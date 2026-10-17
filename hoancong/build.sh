#!/bin/sh
# Dong goi: VERSION (vd 2.1.0) + ma bam src/*.py -> BUILD; ghi src/build.txt; dong goi py/pylibs.zip;
# cap nhat BUILD, tieu de cua so (<title>) va nhan phien ban tren thanh tieu de trong index.html
set -e; cd "$(dirname "$0")"; rm -rf src/__pycache__
V="$(tr -d ' \n' < VERSION)"
B="$V-$(cat src/*.py | sha1sum | cut -c1-6)"
printf '%s\n' "$B" > src/build.txt
rm -f py/pylibs.zip.tmp; cp py/pylibs.zip py/pylibs.zip.tmp
(cd src && zip -q -j ../py/pylibs.zip.tmp *.py build.txt) && mv py/pylibs.zip.tmp py/pylibs.zip
sed -i -e "s/^const BUILD = '[^']*';/const BUILD = '$B';/" \
       -e "s#<title>Hoàn công CSKVTT[^<]*</title>#<title>Hoàn công CSKVTT v$V</title>#" \
       -e "s#<small class=\"appver\" id=\"appver\">[^<]*</small>#<small class=\"appver\" id=\"appver\">v$V</small>#" index.html
echo "VERSION v$V  BUILD $B"
