import json

class Templates(dict):
    def __init__(self):
        super().__init__()
        self.default = "plotly"

templates = Templates()

def to_json(fig, validate=False):
    if hasattr(fig, 'to_dict'):
        return json.dumps(fig.to_dict())
    if isinstance(fig, dict):
        return json.dumps(fig)
    return str(fig)
