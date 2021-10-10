import pandas as pd
data = pd.read_csv('https://www.gstatic.com/covid19/mobility/Global_Mobility_Report.csv')
data = data.loc[data.country_region == 'Australia', :]

chart1 = data.copy()
mask = ((chart1.country_region == 'Australia') 
        & (chart1.sub_region_1 == 'New South Wales') 
        & chart1.sub_region_2.isnull())
chart1 = chart1.loc[mask, :]
chart1['y.0.col'] = 'workplaces_percent_change_from_baseline'
chart1['y.0.name'] = 'Time at Workplace'
chart1['y.1.col'] = 'residential_percent_change_from_baseline'
chart1['y.1.name'] = 'Time at Residence'
chart1['y.1.line_color'] = '000000'
chart1['y.1.line_dash'] = 'dash'
chart1['x_axis.type'] = 'date'
chart1['x_axis.col'] = 'date'
chart1['x_axis.title'] = 'date'
chart1['x_axis.number_format'] = 'dd-MMM'
chart1['x_axis.tick_size'] = 12
chart1['x_axis.tick_position'] = 'low'
chart1['legend.enabled'] = 'true'
chart1['legend.position'] = 'bottom'
chart1['chart.type'] = 'line'
chart1['chart.width'] = 20
chart1['chart.height'] = 10
chart1['chart.title'] = 'Mobility'
chart1['chart.id'] = 'CHART_1'
chart1['y_axis.title'] = 'Mobility'
chart1['y_axis.tick_size'] = 12
chart1.to_csv('chart1.csv', index=False)

chart2 = data.copy()
mask = ((chart2.country_region == 'Australia') 
        & chart2.sub_region_2.isnull())
chart2 = chart2.loc[mask, :]

chart2['y.0.col'] = 'workplaces_percent_change_from_baseline'
chart2['y.0.name'] = 'Time at Workplace'
chart2['y.0.line_color'] = '0000FF'

chart2['y.1.col'] = 'residential_percent_change_from_baseline'
chart2['y.1.name'] = 'Time at Residence'
chart2['y.1.line_color'] = 'FF0000'

chart2['y_axis.tick_size'] = 8

chart2['x_axis.type'] = 'date'
chart2['x_axis.col'] = 'date'
chart2['x_axis.number_format'] = 'dd-MMM'
chart2['x_axis.tick_size'] = 8
chart2['x_axis.tick_position'] = 'low'

chart2['legend.enabled'] = 'false'

chart2['chart.type'] = 'line'
chart2['chart.width'] = 6
chart2['chart.height'] = 5.333
chart2['chart.title'] = chart2['sub_region_1']
chart2['chart.title_size'] = 8

chart2['facet.col'] = 'sub_region_1'
chart2['chart.id'] = 'CHART_2__' + chart2['sub_region_1']

chart2.to_csv('chart2.csv', index=False)

chart3 = data.copy()
chart3['test'] = 20
mask = ((chart3.country_region == 'Australia') 
        & (chart3.sub_region_1 == 'New South Wales') 
        & chart3.sub_region_2.isnull())
chart3 = chart3.loc[mask, :]
chart3['y.0.col'] = 'test'
chart3['y.0.name'] = 'Time at Workplace'
chart3['y.1.col'] = 'grocery_and_pharmacy_percent_change_from_baseline'
chart3['y.1.name'] = 'Time at Residence'
chart3['y.1.line_color'] = '000000'
chart3['y.1.line_dash'] = 'dash'
chart3['x_axis.type'] = 'date'
chart3['x_axis.col'] = 'date'
chart3['x_axis.title'] = 'date'
chart3['x_axis.number_format'] = 'dd-MMM'
chart3['x_axis.tick_size'] = 12
chart3['x_axis.tick_position'] = 'low'
chart3['legend.enabled'] = 'true'
chart3['legend.position'] = 'bottom'
chart3['chart.type'] = 'line'
chart3['chart.width'] = 20
chart3['chart.height'] = 10
chart3['chart.title'] = 'Mobility'
chart3['chart.id'] = 'CHART_1'
chart3['y_axis.title'] = 'Mobility'
chart3['y_axis.tick_size'] = 12
chart3.to_csv('chart3.csv', index=False)


chart4 = data.copy()
chart4['test'] = 20
mask = ((chart4.country_region == 'Australia') 
        & chart4.sub_region_2.isnull())
chart4 = chart4.loc[mask, :]

chart4['y.0.col'] = 'test'
chart4['y.0.name'] = 'Time at Workplace'
chart4['y.0.line_color'] = '0000FF'

chart4['y.1.col'] = 'residential_percent_change_from_baseline'
chart4['y.1.name'] = 'Time at Residence'
chart4['y.1.line_color'] = 'FF0000'

chart4['y_axis.tick_size'] = 8

chart4['x_axis.type'] = 'date'
chart4['x_axis.col'] = 'date'
chart4['x_axis.number_format'] = 'dd-MMM'
chart4['x_axis.tick_size'] = 8
chart4['x_axis.tick_position'] = 'low'

chart4['legend.enabled'] = 'false'

chart4['chart.type'] = 'line'
chart4['chart.width'] = 6
chart4['chart.height'] = 5.333
chart4['chart.title'] = chart4['sub_region_1']
chart4['chart.title_size'] = 8

chart4['facet.col'] = 'sub_region_1'
chart4['chart.id'] = 'CHART_2__' + chart4['sub_region_1']

chart4.to_csv('chart4.csv', index=False)