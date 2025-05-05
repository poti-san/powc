from powcd2d.dxgi import DXGIFactory

factory = DXGIFactory.create()

for adapter in factory.adapter_iter:
    print(adapter.desc.description)
