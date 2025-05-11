# powcwmi

PythonからWindowsのWMIを使いやすくするパッケージです。powcパッケージに依存します。

このパッケージは一部の関数をテストしていません。特にWMIインスタンスのメソッド呼び出しは注意してください。

## ROOT/CIMV2名前空間のWMIクラスを列挙する。

```python
from operator import attrgetter

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
server = locator.connect_server(r"root\cimv2")

for cls in sorted(server.create_classenum(None, shallow=False), key=attrgetter("classname")):
    print(cls.classname)
```

## Win32_ComputerSystemのメソッド名を取得する。

```python
from pprint import pp

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
service = locator.connect_server(r"root\cimv2")

for instance in service.create_instanceenum("Win32_ComputerSystem", shallow=False):
    cls = service.get_object(instance.classname or "")
    pp(cls.methodnames_all)
```

## Win32_ComputerSystemのSystemFamilyプロパティを取得する。

```python
from pprint import pp

from powcwmi import WBEMLocator

locator = WBEMLocator.create()
service = locator.connect_server(r"root\cimv2")

for instance in service.create_instanceenum("Win32_ComputerSystem"):
    pp((instance.get("SystemFamily").value.to_str(), instance.get("SystemSKUNumber").value.to_str()))
```

## ROOT名前空間直下の名前空間名を取得する。

```python
from powcwmi import WBEMLocator

locator = WBEMLocator.create()
server = locator.connect_server("root")

for cls in sorted(server.create_instanceenum("__Namespace"), key=lambda cls: str(cls.get("Name").value)):
    print(cls.get("Name").value)
```
