import sys
from .express import Figure, Layout

class AnyClass:
    def __init__(self, *args, **kwargs):
        pass
    def __call__(self, *args, **kwargs):
        return self
    def __getattr__(self, name):
        return AnyClass()
    def __getitem__(self, key):
        return AnyClass()
    def __setitem__(self, key, val):
        pass

class MagicModule(sys.modules[__name__].__class__):
    def __getattr__(self, name):
        if name in ('Figure', 'Layout'):
            return globals()[name]
        return AnyClass()

sys.modules[__name__].__class__ = MagicModule

__all__ = ['Figure', 'Layout']
