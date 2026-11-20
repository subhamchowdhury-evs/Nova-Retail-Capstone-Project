"""Generate the deck from the workbook, template, icons, and ADK-produced insights."""
from pathlib import Path
import json, openpyxl
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Pt
from pptx.dml.color import RGBColor

BASE=Path(__file__).parent; DATA=BASE/'novaretail_dataset.xlsx'; TEMPLATE=BASE/'novaretail_template.pptx'; OUTPUT=BASE/'NovaRetail_Q2_2026_Executive_Deck.pptx'
NAVY=RGBColor(19,41,75); WHITE=RGBColor(255,255,255); TEAL=RGBColor(31,153,147); COLORS=[RGBColor(38,100,180),RGBColor(31,153,147),RGBColor(46,139,87),RGBColor(195,63,63)]
def rows(ws):
 r=list(ws.iter_rows(values_only=True)); return [dict(zip(r[0],x)) for x in r[1:] if any(v is not None for v in x)]
def shape(s,n): return next(x for x in s.shapes if x.name==n)
def delete(x): x._element.getparent().remove(x._element)
def text(x,value,size=16,color=NAVY,bold=False,bullets=False):
 tf=x.text_frame; tf.clear(); tf.word_wrap=True
 for i,line in enumerate(value if isinstance(value,list) else [value]):
  p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.text=('• ' if bullets else '')+line; p.font.name='Aptos'; p.font.size=Pt(size); p.font.color.rgb=color; p.font.bold=bold; p.space_after=Pt(8 if bullets else 0)
def icon(s,n,file):
 ph=shape(s,n); l,t,w,h=ph.left,ph.top,ph.width,ph.height; delete(ph)
 try: delete(shape(s,n+'_label'))
 except StopIteration: pass
 s.shapes.add_picture(str(BASE/file),l,t,width=w,height=h)
def chart(s,box,label,kind,cats,series,title):
 ph=shape(s,box); l,t,w,h=ph.left,ph.top,ph.width,ph.height; delete(ph); delete(shape(s,label)); d=ChartData(); d.categories=cats
 for name,vals in series: d.add_series(name,vals)
 c=s.shapes.add_chart(kind,l,t,w,h,d).chart; c.has_legend=len(series)>1;
 if c.has_legend: c.legend.position=XL_LEGEND_POSITION.BOTTOM; c.value_axis.has_title=True; c.value_axis.axis_title.text_frame.text=title
 for i,ser in enumerate(c.series): ser.format.fill.solid(); ser.format.fill.fore_color.rgb=COLORS[i]; ser.format.line.color.rgb=COLORS[i]
 return c

def build_deck(insights):
 wb=openpyxl.load_workbook(DATA,data_only=True); perf=rows(wb['Regional_Performance']); cx=rows(wb['Customer_Experience']); risk=rows(wb['Operational_Risk'])
 q=list(dict.fromkeys(x['Quarter'] for x in perf)); regs=list(dict.fromkeys(x['Region'] for x in perf)); latest,prev=q[-1],q[-2]
 total=sum(x['Revenue_USD_M'] for x in perf if x['Quarter']==latest); prior=sum(x['Revenue_USD_M'] for x in perf if x['Quarter']==prev); growth=total/prior-1
 rev={x['Region']:x['Revenue_USD_M'] for x in perf if x['Quarter']==latest}; nps={x['Region']:x['NPS'] for x in cx if x['Quarter']==latest}; wnps=sum(rev[r]*nps[r] for r in regs)/total
 prs=Presentation(TEMPLATE)
 s=prs.slides[0]; text(shape(s,'subtitle_1'),'Q2 2026 • Executive Review',14,WHITE); vals=[(f'${total/1000:.2f}B','Q2 2026 Revenue'),(f'{growth:+.1%}','Quarter-over-Quarter Growth'),(f'{wnps:.1f}','Revenue-Weighted NPS')]
 for i,(v,l) in enumerate(vals,1): text(shape(s,f'kpi{i}_value'),v,28,TEAL,True); text(shape(s,f'kpi{i}_label'),l,12,WHITE)
 text(shape(s,'summary_text'),insights['executive_summary'],17,WHITE); icon(s,'icon_1','growth_chart.png')
 s=prs.slides[1]; text(shape(s,'title_2'),insights['performance_headline'],26,NAVY,True); series=[(r,[next(x['Revenue_USD_M'] for x in perf if x['Region']==r and x['Quarter']==qq) for qq in q]) for r in regs]; chart(s,'chart_performance','chart_performance_label',XL_CHART_TYPE.COLUMN_CLUSTERED,q,series,'Revenue ($M)'); text(shape(s,'insights_performance'),insights['performance_insights'],15,NAVY,False,True); icon(s,'icon_2','growth_chart.png')
 s=prs.slides[2]; text(shape(s,'title_3'),insights['customer_headline'],26,NAVY,True); series=[(r,[next(x['NPS'] for x in cx if x['Region']==r and x['Quarter']==qq) for qq in q]) for r in regs]; chart(s,'chart_customer','chart_customer_label',XL_CHART_TYPE.LINE_MARKERS,q,series,'NPS'); text(shape(s,'insights_customer'),insights['customer_insights'],15,NAVY,False,True); icon(s,'icon_3','customer_star.png')
 s=prs.slides[3]; text(shape(s,'title_4'),insights['risk_headline'],26,NAVY,True); lr=[x for x in risk if x['Quarter']==latest]; c=chart(s,'chart_risk','chart_risk_label',XL_CHART_TYPE.BAR_CLUSTERED,[x['Region'] for x in lr],[('Stockout rate',[x['Stockout_Rate_pct']/100 for x in lr])],'Stockout rate'); c.has_legend=False; c.value_axis.tick_labels.number_format='0.0%'; text(shape(s,'insights_risk'),insights['risk_insights'],15,NAVY,False,True); icon(s,'icon_4','risk_alert.png')
 s=prs.slides[4]; text(shape(s,'recommendations_text'),insights['recommendations'],17,WHITE,False,True); text(shape(s,'closing_statement'),insights['closing_statement'],16,WHITE,True); icon(s,'icon_5','lightbulb_action.png')
 prs.save(OUTPUT)

if __name__=='__main__':
 with open(BASE/'generated_insights.json',encoding='utf-8') as f: build_deck(json.load(f))
 print(f'Created: {OUTPUT}')
