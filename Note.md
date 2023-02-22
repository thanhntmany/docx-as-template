Ubuntu:

Setup code:

```bash
sudo apt-get install libreoffice --no-install-recommends
sudo apt install libreoffice-java-common
rm ~/.config/libreoffice/4/user/config/javasettings_Linux_*.xml
sudo apt install default-jre
```

Cover code:
```bash
libreoffice --headless --convert-to pdf MyWordFile.docx --outdir ./
```
