build_exec:
	pyinstaller \
	--onefile \
	-w \
	--name "excel-merge" \
	--exclude-module=autopep8 \
	--exclude-module=numpy \
	--exclude-module=pycodestyle \
	--exclude-module=pytz \
	--exclude-module=six \
	--exclude-module=tomli \
	--exclude-module=altgraph \
	--exclude-module=et-xmlfile \
	--exclude-module=macholib \
	--exclude-module=pyinstaller \
	--exclude-module=pyinstaller-hooks-contrib \
	--exclude-module=python-dateutil \
	'main.py'