from pathlib import Path
from copy import deepcopy
import openpyxl
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt

BASE = Path(__file__).resolve().parent
TEMPLATE = BASE / 'novaretail_template.pptx'
DATASET = BASE / 'novaretail_dataset.xlsx'
OUTPUT = BASE / 'NovaRetail_Q2_2026_Executive_Deck.pptx'

NAVY = RGBColor(19, 41, 75)
BLUE = RGBColor(38, 100, 180)
TEAL = RGBColor(31, 153, 147)
GREEN = RGBColor(46, 139, 87)
AMBER = RGBColor(238, 167, 52)
RED = RGBColor(195, 63, 63)
MID = RGBColor(91, 104, 122)
LIGHT = RGBColor(224, 230, 238)
WHITE = RGBColor(255,255,255)
REGION_COLORS = [BLUE, TEAL, GREEN, RED]


def read_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    return [dict(zip(headers, r)) for r in rows[1:] if any(v is not None for v in r)]


def delete_shape(shape):
    sp = shape._element
    sp.getparent().remove(sp)


def shape_by_name(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise KeyError(f'Missing expected template shape: {name}')


def set_text(shape, text, size=18, color=NAVY, bold=False, bullets=False, line_spacing=1.05):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Pt(7)
    tf.margin_right = Pt(7)
    tf.margin_top = Pt(3)
    tf.margin_bottom = Pt(3)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.name = 'Aptos'
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.space_after = Pt(8 if bullets else 0)
        p.line_spacing = line_spacing
        if bullets:
            p.level = 0
            p.text = '• ' + line
    return shape


def style_title(shape, text):
    set_text(shape, text, size=26, color=NAVY, bold=True)
    shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def add_icon(slide, placeholder_name, label_name, symbol, color):
    ph = shape_by_name(slide, placeholder_name)
    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    delete_shape(ph)
    try:
        delete_shape(shape_by_name(slide, label_name))
    except KeyError:
        pass
    size = min(width, height)
    x = left + (width-size)//2
    y = top + (height-size)//2
    icon = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    icon.fill.solid(); icon.fill.fore_color.rgb = color
    icon.line.color.rgb = color
    tf = icon.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.text = symbol; p.alignment = PP_ALIGN.CENTER
    p.font.name = 'Arial'; p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = WHITE
    return icon


def add_chart(slide, placeholder_name, label_name, chart_type, categories, series, colors, value_title=None, percent=False):
    ph = shape_by_name(slide, placeholder_name)
    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    delete_shape(ph)
    delete_shape(shape_by_name(slide, label_name))
    data = ChartData(); data.categories = categories
    for name, values in series:
        data.add_series(name, values)
    chart = slide.shapes.add_chart(chart_type, left, top, width, height, data).chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.name = 'Aptos'; chart.legend.font.size = Pt(10)
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = LIGHT
    chart.value_axis.tick_labels.font.name = 'Aptos'; chart.value_axis.tick_labels.font.size = Pt(10)
    chart.category_axis.tick_labels.font.name = 'Aptos'; chart.category_axis.tick_labels.font.size = Pt(10)
    if value_title:
        chart.value_axis.has_title = True; chart.value_axis.axis_title.text_frame.text = value_title
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.name = 'Aptos'
        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(10)
    if percent:
        chart.value_axis.tick_labels.number_format = '0.0%'
    for idx, ser in enumerate(chart.series):
        ser.format.line.color.rgb = colors[idx]
        ser.format.line.width = Pt(2.25)
        ser.format.fill.solid(); ser.format.fill.fore_color.rgb = colors[idx]
        if chart_type == XL_CHART_TYPE.LINE_MARKERS:
            ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = colors[idx]
            ser.marker.format.line.color.rgb = colors[idx]
    return chart


def main():
    wb = openpyxl.load_workbook(DATASET, data_only=True)
    perf = read_sheet(wb['Regional_Performance'])
    cx = read_sheet(wb['Customer_Experience'])
    risk = read_sheet(wb['Operational_Risk'])
    quarters = list(dict.fromkeys(r['Quarter'] for r in perf))
    regions = list(dict.fromkeys(r['Region'] for r in perf))
    latest, previous = quarters[-1], quarters[-2]

    total_latest = sum(r['Revenue_USD_M'] for r in perf if r['Quarter']==latest)
    total_previous = sum(r['Revenue_USD_M'] for r in perf if r['Quarter']==previous)
    qoq = total_latest/total_previous - 1
    latest_rev = {r['Region']: r['Revenue_USD_M'] for r in perf if r['Quarter']==latest}
    latest_nps = {r['Region']: r['NPS'] for r in cx if r['Quarter']==latest}
    weighted_nps = sum(latest_rev[x]*latest_nps[x] for x in regions)/total_latest

    prs = Presentation(TEMPLATE)

    # Slide 1
    s = prs.slides[0]
    set_text(shape_by_name(s,'subtitle_1'), 'Q2 2026 • Executive Review', 15, MID, False)
    for val, lab, vn, ln in [
        (f'${total_latest/1000:.2f}B','Q2 2026 Revenue','kpi1_value','kpi1_label'),
        (f'{qoq:+.1%}','Quarter-over-Quarter Growth','kpi2_value','kpi2_label'),
        (f'{weighted_nps:.1f}','Revenue-Weighted NPS','kpi3_value','kpi3_label')]:
        set_text(shape_by_name(s,vn),val,28,TEAL,True)
        set_text(shape_by_name(s,ln),lab,12,WHITE,False)
    summary = ('NovaRetail closed Q2 2026 with $1.28B in revenue, up 4.5% quarter over quarter. '
               'North America remains the largest market, while APAC is the clearest growth engine, combining a 27% revenue increase since Q3 2025 with the highest NPS. '
               'Latin America is the priority intervention: revenue contracted as returns, delivery delays, stockouts and fraud all worsened.')
    set_text(shape_by_name(s,'summary_text'),summary,17,WHITE,False,line_spacing=1.12)
    add_icon(s,'icon_1','icon_1_label','↗',TEAL)

    # Slide 2
    s = prs.slides[1]
    style_title(shape_by_name(s,'title_2'),'APAC Accelerates as Latin America Contracts')
    perf_series=[]
    for reg in regions:
        perf_series.append((reg,[next(r['Revenue_USD_M'] for r in perf if r['Region']==reg and r['Quarter']==q) for q in quarters]))
    add_chart(s,'chart_performance','chart_performance_label',XL_CHART_TYPE.COLUMN_CLUSTERED,quarters,perf_series,REGION_COLORS,'Revenue ($M)')
    bullets=[
        'APAC revenue rose from $260M to $330M, a 27% increase across the four-quarter period.',
        'North America remains the scale leader at $465M in Q2, despite uneven quarterly momentum.',
        'Europe stabilized at $345M, above Q3 2025 but still below its Q4 peak.',
        'Latin America fell to $140M while its returns rate reached 8.2%, the highest in the portfolio.'
    ]
    set_text(shape_by_name(s,'insights_performance'),bullets,15,NAVY,False,True,1.05)
    add_icon(s,'icon_2','icon_2_label','↗',BLUE)

    # Slide 3
    s = prs.slides[2]
    style_title(shape_by_name(s,'title_3'),'Customer Experience Gains Mask a Latin America Breakdown')
    cx_series=[]
    for reg in regions:
        cx_series.append((reg,[next(r['NPS'] for r in cx if r['Region']==reg and r['Quarter']==q) for q in quarters]))
    add_chart(s,'chart_customer','chart_customer_label',XL_CHART_TYPE.LINE_MARKERS,quarters,cx_series,REGION_COLORS,'NPS')
    bullets=[
        'APAC leads with NPS 62 and 92% CSAT, supported by the fastest delivery time at 2.2 days.',
        'North America and Europe improved to NPS 55 and 53 respectively.',
        'Latin America NPS fell from 38 to 29 as average delivery time increased to 7.5 days.',
        'Support tickets in Latin America climbed from 20K to 31K, signaling mounting service friction.'
    ]
    set_text(shape_by_name(s,'insights_customer'),bullets,15,NAVY,False,True,1.05)
    add_icon(s,'icon_3','icon_3_label','★',TEAL)

    # Slide 4
    s = prs.slides[3]
    style_title(shape_by_name(s,'title_4'),'Stockouts and Capacity Pressure Concentrate Risk')
    latest_risk=[r for r in risk if r['Quarter']==latest]
    categories=[r['Region'] for r in latest_risk]
    series=[('Stockout rate',[r['Stockout_Rate_pct']/100 for r in latest_risk])]
    chart=add_chart(s,'chart_risk','chart_risk_label',XL_CHART_TYPE.BAR_CLUSTERED,categories,series,[RED],'Stockout rate',True)
    chart.has_legend=False
    chart.value_axis.maximum_scale=0.14
    chart.value_axis.minimum_scale=0
    chart.plots[0].has_data_labels=True
    chart.plots[0].data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
    chart.plots[0].data_labels.number_format='0.0%'
    chart.plots[0].data_labels.font.name='Aptos'; chart.plots[0].data_labels.font.size=Pt(10)
    bullets=[
        'Latin America is the critical risk cluster: 12.5% stockouts, 10-day returns processing and 30 fraud incidents.',
        'APAC warehouse utilization reached 92%, creating a capacity constraint behind its growth story.',
        'APAC stockouts increased to 8.5%, more than double North America and Europe.',
        'North America and Europe remain comparatively controlled, with stockout rates at or below 4.0%.'
    ]
    set_text(shape_by_name(s,'insights_risk'),bullets,15,NAVY,False,True,1.05)
    add_icon(s,'icon_4','icon_4_label','!',RED)

    # Slide 5
    s = prs.slides[4]
    recs=[
       'EXPAND APAC CAPACITY • Add near-term overflow capacity and approve a durable expansion plan before Q4.',
       'STABILIZE LATIN AMERICA • Launch a 90-day service recovery sprint focused on delivery reliability, inventory availability and ticket resolution.',
       'TIGHTEN FRAUD CONTROLS • Increase monitoring and exception review in Latin America, where incidents rose to 30.',
       'REDUCE RETURNS FRICTION • Diagnose Latin America’s 8.2% return rate and cut processing time from 10 days with clearer ownership and targets.'
    ]
    set_text(shape_by_name(s,'recommendations_text'),recs,17,WHITE,False,True,1.10)
    set_text(shape_by_name(s,'closing_statement'),'Protect the APAC growth engine while making Latin America recovery the next-quarter leadership priority.',17,WHITE,True)
    shape_by_name(s,'closing_statement').fill.solid(); shape_by_name(s,'closing_statement').fill.fore_color.rgb=NAVY
    add_icon(s,'icon_5','icon_5_label','✓',GREEN)

    # Update footers and remove template suffix
    for idx, slide in enumerate(prs.slides,1):
        fn=f'footer_{idx}'
        try: set_text(shape_by_name(slide,fn),'NovaRetail Quarterly Business Review',9,MID,False)
        except KeyError: pass

    prs.save(OUTPUT)
    print(f'Created {OUTPUT}')

if __name__ == '__main__':
    main()
