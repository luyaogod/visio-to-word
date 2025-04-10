打包命令
1. 无图标
```
pyinstaller --onefile --name=V2WTools script.py
```

2. 有图标
```
pyinstaller --onefile --name=V2WTools --icon=你的图标.ico script.py
```

3. 无图标-无窗口
```
pyinstaller --onefile --name=V2WTools --noconsole script.py
```
