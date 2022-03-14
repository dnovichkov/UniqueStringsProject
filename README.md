Работает все это с третьей версией Python.

Рекомендуется настроить виртуальное окружение. Вот "правильные" ссылки:
[https://docs.python.org/3/library/venv.html](https://docs.python.org/3/library/venv.html) и 
[https://virtualenv.pypa.io/en/stable/userguide/](https://docs.python.org/3/library/venv.html)
На своей Win-машине я сделал это так:

```
python -m venv env
call env/scripts/activate.bat
```

После этого устанавливаем необходимые модули

```
pip install -r requirements.txt
```

Программу необходимо создать в виде исполняемого файла для дальнейшего запуска в виде службы. Делается это из директории проекта запуском команды:
 ```
 pyinstaller main.spec
  ```

 
После этого в директории *dist\main* запустить командную строку и выполнить следующую команду:
```
main.exe
```

Альтернативные варианты:
```
main.exe -h

main.exe -dest="RES2.xlsx" -src=2.xlsx

main.exe -src="2.xlsx"

```
