# powc

PythonからCOMの基本機能を使いやすくするパッケージです。comtypesパッケージに依存します。 

**システムプロパティの列挙**

```python
from powcpropsys.propsys import PropertySystem

ps = PropertySystem.create()

for desc in ps.propdescs_all:
    print(f"{str(desc)} {desc.propkey}")
```
