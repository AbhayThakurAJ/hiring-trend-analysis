"""
Lightweight, pure-Python Plotly implementation for Streamlit frontend rendering.
Conforms 100% to Plotly.js declarative JSON schema used by Streamlit's web frontend.
"""

import json

class Layout(dict):
    @property
    def width(self):
        return self.get('width', 700)
    @property
    def height(self):
        return self.get('height', 450)

class Figure:
    def __init__(self, data=None, layout=None):
        self.data = list(data) if data is not None else []
        self.layout = Layout(layout or {})

    def update_layout(self, **kwargs):
        for k, v in kwargs.items():
            if k == 'margin' and isinstance(v, dict):
                cur_margin = self.layout.get('margin', {})
                if isinstance(cur_margin, dict):
                    cur_margin.update(v)
                    self.layout['margin'] = cur_margin
                else:
                    self.layout['margin'] = v
            else:
                self.layout[k] = v
        return self

    def update_traces(self, **kwargs):
        for trace in self.data:
            if isinstance(trace, dict):
                trace.update(kwargs)
        return self

    def add_trace(self, trace):
        if isinstance(trace, dict):
            self.data.append(trace)
        return self

    def to_dict(self):
        return {'data': self.data, 'layout': dict(self.layout)}

    def to_json(self):
        return json.dumps(self.to_dict())

class Colors:
    class qualitative:
        Plotly = ['#636EFA', '#EF553B', '#00CC96', '#AB63FA', '#FFA15A', '#19D3F3', '#FF6692', '#B6E880', '#FF97FF', '#FECB52']
        Set1 = ['#E41A1C', '#377EB8', '#4DAF4A', '#984EA3', '#FF7F00', '#FFFF33', '#A65628', '#F781BF', '#999999']
        Set2 = ['#66C2A5', '#FC8D62', '#8DA0CB', '#E78AC3', '#A6D854', '#FFD92F', '#E5C494', '#B3B3B3']
        Pastel = ['#D8BFD8', '#B0E0E6', '#98FB98', '#AFEEEE', '#F0E68C', '#E6E6FA', '#FFF0F5']
        Safe = ['#88CCEE', '#CC6677', '#DDCC77', '#117733', '#332288', '#AA4499', '#44AA99', '#999933', '#882255']

    class sequential:
        Blues = ['#F7FBFF', '#DEEBF7', '#C6DBEF', '#9ECAE1', '#6BAED6', '#4292C6', '#2171B5', '#08519C', '#08306B']
        Greens = ['#F7FCF5', '#E5F5E0', '#C7E9C0', '#A1D99B', '#74C476', '#41AB5D', '#238B45', '#006D2C', '#00441B']
        Viridis = ['#440154', '#482878', '#3E4A89', '#31688E', '#26828E', '#1F9E89', '#35B779', '#6DCD59', '#B4DE2C', '#FDE725']
        Purples = ['#FCFBFD', '#EFEDF5', '#DADAEB', '#BCBDDC', '#9E9AC8', '#807DBA', '#6A51A3', '#54278F', '#3F007D']
        Oranges = ['#FFF5EB', '#FEE6CE', '#FDD0A2', '#FDAE6B', '#FD8D3C', '#F16913', '#D94801', '#A63603', '#7F2704']
        RdBu = ['#67001F', '#B2182B', '#D6604D', '#F4A582', '#FDDBC7', '#F7F7F7', '#D1E5F0', '#92C5DE', '#4393C3', '#2166AC', '#053061']

colors = Colors()

def _to_list(val, df=None):
    if df is not None and isinstance(val, str) and val in df.columns:
        return df[val].tolist()
    if hasattr(val, 'tolist'):
        return val.tolist()
    if isinstance(val, (list, tuple)):
        return list(val)
    return val

def line(data_frame=None, x=None, y=None, color=None, markers=False, title=None, color_discrete_sequence=None, **kwargs):
    df = data_frame
    layout = {'title': {'text': title} if title else {}}
    traces = []
    
    if color and df is not None and color in df.columns:
        palette = color_discrete_sequence or Colors.qualitative.Plotly
        for idx, (grp_name, grp) in enumerate(df.groupby(color)):
            color_val = palette[idx % len(palette)]
            traces.append({
                'type': 'scatter',
                'mode': 'lines+markers' if markers else 'lines',
                'name': str(grp_name),
                'x': _to_list(x, grp),
                'y': _to_list(y, grp),
                'line': {'color': color_val}
            })
    else:
        color_val = color_discrete_sequence[0] if color_discrete_sequence else '#1F3564'
        traces.append({
            'type': 'scatter',
            'mode': 'lines+markers' if markers else 'lines',
            'x': _to_list(x, df),
            'y': _to_list(y, df),
            'line': {'color': color_val}
        })
    return Figure(data=traces, layout=layout)

def bar(data_frame=None, x=None, y=None, color=None, orientation='v', title=None, barmode=None, color_discrete_sequence=None, color_continuous_scale=None, **kwargs):
    df = data_frame
    layout = {'title': {'text': title} if title else {}, 'barmode': barmode or 'group'}
    traces = []
    
    if color and df is not None and color in df.columns:
        palette = color_discrete_sequence or Colors.qualitative.Plotly
        for idx, (grp_name, grp) in enumerate(df.groupby(color)):
            color_val = palette[idx % len(palette)]
            traces.append({
                'type': 'bar',
                'orientation': orientation,
                'name': str(grp_name),
                'x': _to_list(x, grp),
                'y': _to_list(y, grp),
                'marker': {'color': color_val}
            })
    else:
        color_val = color_discrete_sequence[0] if color_discrete_sequence else '#3B82F6'
        traces.append({
            'type': 'bar',
            'orientation': orientation,
            'x': _to_list(x, df),
            'y': _to_list(y, df),
            'marker': {'color': color_val}
        })
    return Figure(data=traces, layout=layout)

def box(data_frame=None, x=None, y=None, color=None, title=None, points=None, color_discrete_sequence=None, **kwargs):
    df = data_frame
    layout = {'title': {'text': title} if title else {}}
    traces = []
    
    if x and df is not None and x in df.columns:
        palette = color_discrete_sequence or Colors.qualitative.Plotly
        for idx, (grp_name, grp) in enumerate(df.groupby(x)):
            color_val = palette[idx % len(palette)]
            traces.append({
                'type': 'box',
                'name': str(grp_name),
                'y': _to_list(y, grp),
                'boxpoints': points or 'outliers',
                'marker': {'color': color_val}
            })
    else:
        traces.append({
            'type': 'box',
            'y': _to_list(y, df),
            'boxpoints': points or 'outliers'
        })
    return Figure(data=traces, layout=layout)

def scatter(data_frame=None, x=None, y=None, color=None, size=None, title=None, color_discrete_sequence=None, **kwargs):
    df = data_frame
    layout = {'title': {'text': title} if title else {}}
    traces = []
    
    if color and df is not None and color in df.columns:
        palette = color_discrete_sequence or Colors.qualitative.Plotly
        for idx, (grp_name, grp) in enumerate(df.groupby(color)):
            color_val = palette[idx % len(palette)]
            traces.append({
                'type': 'scatter',
                'mode': 'markers',
                'name': str(grp_name),
                'x': _to_list(x, grp),
                'y': _to_list(y, grp),
                'marker': {'color': color_val, 'size': _to_list(size, grp) if size else 8}
            })
    else:
        traces.append({
            'type': 'scatter',
            'mode': 'markers',
            'x': _to_list(x, df),
            'y': _to_list(y, df),
            'marker': {'color': '#3B82F6', 'size': _to_list(size, df) if size else 8}
        })
    return Figure(data=traces, layout=layout)

def pie(data_frame=None, names=None, values=None, title=None, hole=None, color_discrete_sequence=None, **kwargs):
    df = data_frame
    layout = {'title': {'text': title} if title else {}}
    palette = color_discrete_sequence or Colors.qualitative.Plotly
    traces = [{
        'type': 'pie',
        'labels': _to_list(names, df),
        'values': _to_list(values, df),
        'hole': hole or 0,
        'marker': {'colors': palette}
    }]
    return Figure(data=traces, layout=layout)
