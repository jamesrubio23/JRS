"""
Spyder Editor

This is a temporary script file.

http://www.chadfulton.com/topics/statespace_large_dynamic_factor_models.html

https://www.statsmodels.org/stable/generated/statsmodels.tsa.statespace.dynamic_factor_mq.DynamicFactorMQ.html
https://www.statsmodels.org/devel/examples/notebooks/generated/statespace_news.html

https://www.sr-sv.com/nowcasting-for-financial-markets/
"""
import numpy as np
import pandas as pd
import statsmodels.api as sm
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw
import time
import gc
start_time = time.time()

def transform(column, transforms):
    transformation = transforms[column.name]
    mult = 1
    if transformation == 1:
        pass
    elif transformation == 2:
        column = column.diff()
    elif transformation == 3:
        column = column.diff().diff()
    elif transformation == 4:
        column = np.log(column)
    elif transformation == 5:
        column = np.log(column).diff() * 100 * mult
    elif transformation == 6:
        column = np.log(column).diff().diff() * 100 * mult
    elif transformation == 7:
        column = ((column / column.shift(1)) ** mult - 1.0) * 100
    elif transformation == 8:
        column = ((column / column.shift(12)) ** mult - 1.0) * 100
    elif transformation == 11:
        column = column.shift(1)
    elif transformation == 12:
        column = column.shift(2)
    elif transformation == 13:
        column = column.shift(3)
    elif transformation == 14:
        column = column.shift(4)
    elif transformation == 15:
        column = column.shift(5)
    elif transformation == 112:
        column = column.shift(12)
    return column
    gc.collect()

def remove_outliers(dta):
    mean = dta.mean()
    iqr = dta.quantile([0.25, 0.75]).diff().T.iloc[:, 1]
    mask = np.abs(dta) > mean + 10 * iqr
    treated = dta.copy()
    treated[mask] = np.nan
    return treated
    gc.collect()
'\n####################\n### LOADING DATA ###\n####################\n'
path = 'Y:/mesa teso/JRS/8_DFM Nowcasting_US/'
path += 'Datos_US_large.xlsx'
Book = xw.Book(path)
Dibujar_graficos = Book.sheets('Indice').range('c5').options(index=False).value
Vintages = Book.sheets('Indice').range('c6').options(expand='right').value
" \nfecha_hoy =  datetime.today().strftime('%d-%m-%Y') # Nueva hoja con la vintage de hoy\nVintages = Vintages+[fecha_hoy]\nBook.sheets('Indice').range('c6').value = Vintages #añado la vintage de hoy a la lista de vintages\n"
Vintage_0 = Vintages[0]
Vintage_last = Vintages[-1]
Datos_vinculados = Book.sheets('Datos').range('A1').options(pd.DataFrame, expand='table', index=False, decimal='.').value
try:
    Book.sheets.add(Vintage_last)
except:
    print('Hoja ya creada')
Book.sheets(Vintage_last).range('a1').options(index=False).value = Datos_vinculados

def load_fredmd_data(vintage):
    orig_m = Book.sheets(vintage).range('A1').options(pd.DataFrame, expand='table', index=False, decimal='.').value
    orig_m = orig_m.drop(columns=['SERIES'])
    orig_m = orig_m.drop(columns=['GDP q/q (SAAR)'])
    transform_m = orig_m.iloc[1, 1:]
    factor_m1 = orig_m.iloc[2, 1:]
    factor_m2 = orig_m.iloc[3, 1:]
    orig_m = orig_m.iloc[4:]
    orig_m.index = pd.PeriodIndex(orig_m.fecha.tolist(), freq='M')
    orig_m = orig_m.iloc[:, 1:]
    for i in orig_m.columns:
        orig_m[i] = pd.to_numeric(orig_m[i], errors='coerce')
    dta_m = orig_m.apply(transform, axis=0, transforms=transform_m)
    dta_m.loc[:'2019-12'] = remove_outliers(dta_m.loc[:'2019-12'])
    orig_q = Book.sheets(vintage).range('A1').options(pd.DataFrame, expand='table', index=False, decimal='.').value
    orig_q = orig_q[['fecha', 'GDP q/q (SAAR)']]
    orig_q = orig_q.dropna()
    transform_q = orig_q.iloc[1, 1:]
    factor_q1 = orig_q.iloc[2, 1:]
    factor_q2 = orig_q.iloc[3, 1:]
    orig_q = orig_q.iloc[4:]
    orig_q.index = pd.PeriodIndex(orig_q.fecha.tolist(), freq='Q')
    orig_q = orig_q.iloc[:, 1:]
    for i in orig_q.columns:
        orig_q[i] = pd.to_numeric(orig_q[i], errors='coerce')
    dta_q = orig_q.apply(transform, axis=0, transforms=transform_q)
    dta_q.loc[:'2019-12'] = remove_outliers(dta_q.loc[:'2019-12'])
    return types.SimpleNamespace(orig_m=orig_m, orig_q=orig_q, dta_m=dta_m, transform_m=transform_m, factor_m2=factor_m2, dta_q=dta_q, transform_q=transform_q, factor_q2=factor_q2)
    gc.collect()
dta = {date: load_fredmd_data(date) for date in Vintages}
endog_qmonthly = dta[Vintage_last].dta_q
endog_qmonthly = endog_qmonthly.asfreq('M')
Muestra_transformada = dta[Vintage_last].dta_m.join(endog_qmonthly, lsuffix='_')
Book.sheets(Vintage_last).range('Ga5').options(index=False).value = Muestra_transformada
groups = dta[Vintage_0].factor_m2.copy()
groups = groups.append(dta[Vintage_0].factor_q2)
groups = pd.DataFrame(groups).reset_index()
groups.columns = ['description', 'group']
if Dibujar_graficos == 1:
    groups.groupby('group', sort=False).count().rename({'description': '# series in group'}, axis=1)
    factors = {row['description']: ['Global', row['group']] for ix, row in groups.iterrows()}
    print(factors['ISM MANUF'])
factor_multiplicities = {'Global': 1}
factor_multiplicities = 1
factor_orders = {('Hard', 'Soft'): 1, 'Global': 2}
factor_orders = 2
'\n####################\n###    MODELO    ###\n####################\n'
start = '1996'
endog_m = dta[Vintage_0].dta_m.loc[start:, :]
endog_q = dta[Vintage_0].dta_q.loc[start:, :]
'\nmodel = sm.tsa.DynamicFactorMQ(\n    endog_m, endog_quarterly=endog_q,\n    factors=factors, factor_orders=factor_orders,\n    factor_multiplicities=factor_multiplicities)\n'
model = sm.tsa.DynamicFactorMQ(endog_m, endog_quarterly=endog_q, factors=1, factor_orders=1, factor_multiplicities=1)
"\nmodel = sm.tsa.DynamicFactorMQ(\n    endog_m, endog_quarterly=endog_q,\n    factors=3, factor_orders=1,\n    factor_multiplicities=1, idiosyncratic_ar1=True, error_order=1, trend='c')\n"
model.summary()
results = model.fit(disp=10)
print(results.summary(display_diagnostics=True))
results.plot_diagnostics('GDP q/q (SAAR)')
print('Log-Likelihood:', results.llf)
print('AIC:', results.aic)
print('BIC:', results.bic)
print('HQIC:', results.hqic)
rsquared = results.get_coefficients_of_determination(method='joint')
print('R² individual:')
print(rsquared)
'\n####################\n### Forecasting ###\n####################\n'
metodo_fore = ['predicted', 'filtered']
celdas_inicio = ['ka65', 'ca65']
for metodo, celda in zip(metodo_fore, celdas_inicio):
    print(metodo)
    print('FORECASTING en cada Vintage')
    prediction_results = results.get_prediction(start='2000', end='2027', information_set=metodo)
    point_predictions = prediction_results.predicted_mean
    Book.sheets(Vintage_0).range(celda).options(index=True).value = point_predictions
    factor_v = {}
    vintage_results = {Vintage_0: results}
    for vintage in Vintages:
        print(vintage)
        updated_endog_m = dta[vintage].dta_m.loc[start:, :]
        updated_endog_q = dta[vintage].dta_q.loc[start:, :]
        vintage_results[vintage] = results.apply(updated_endog_m, endog_quarterly=updated_endog_q)
        prediction_results = vintage_results[vintage].get_prediction(start='2000', end='2027', information_set=metodo)
        point_predictions = prediction_results.predicted_mean
        Book.sheets(vintage).range(celda).options(index=True).value = point_predictions
'\n####################\n####### NEWS #######\n####################\n'
print('NEWS en cada Vintage y para 3 fechas especificas: Q+1, Q+2 y Q+3')
Variable_impactada = 'GDP q/q (SAAR)'
Fecha_impacto1 = Book.sheets('Impactos').range('D3').options(index=False).value
Fecha_impacto2 = Book.sheets('Impactos').range('D71').options(index=False).value
Fecha_impacto3 = Book.sheets('Impactos').range('D139').options(index=False).value
fechas_impacto = [Fecha_impacto1, Fecha_impacto2, Fecha_impacto3]
celdas_inicio = ['Eb5', 'Eo5', 'Fb5']
rangos_limpiar = ['Eb5:Ek60', 'Eo5:Ex60', 'Fb5:Fek60']
for fecha, celda, rango in zip(fechas_impacto, celdas_inicio, rangos_limpiar):
    print(fecha)
    for i in range(1, len(Vintages)):
        print(Vintages[i])
        vintage = Vintages[i]
        prev_vintage = Vintages[i - 1]
        news = vintage_results[vintage].news(vintage_results[prev_vintage], impact_date=fecha, impacted_variable=Variable_impactada, comparison_type='previous')
        details = news.details_by_impact
        details['absolute impact'] = np.abs(details['impact'])
        details = details.sort_values('weight', ascending=False)
        Book.sheets(vintage).range(rango).clear_contents()
        Book.sheets(vintage).range(celda).options(index=True).value = details
'\n####################\n### POST ESTIMACIÓN ###\n####################\n'
if Dibujar_graficos == 1:
    method = 'individual'
    method = 'joint'
    method = 'cumulative'
    rsquared = results.get_coefficients_of_determination(method='individual')
    top_ten = []
    for factor_name in rsquared.columns[:1]:
        top_factor = rsquared[factor_name].sort_values(ascending=False).iloc[:10].round(2).reset_index()
        top_factor.columns = pd.MultiIndex.from_product([[f'Top ten variables explained by {factor_name}'], ['Variable', '$R^2$']])
        top_ten.append(top_factor)
    pd.concat(top_ten, axis=1)
    with sns.color_palette('deep'):
        fig = results.plot_coefficients_of_determination(method='individual', figsize=(14, 9))
        fig.suptitle('$R^2$ - regression on individual factors', fontsize=14, fontweight=600)
        fig.tight_layout(rect=[0, 0, 1, 0.95])
    group_counts = groups.groupby('group', sort=False).count()['description'].cumsum()
    with sns.color_palette('deep'):
        fig = results.plot_coefficients_of_determination(method='joint', figsize=(14, 3))
        ax = fig.axes[0]
        ax.set_ylim(0, 1.2)
        for i in np.arange(1, len(group_counts), 2):
            start = 0 if i == 0 else group_counts[i - 1]
            end = group_counts[i] + 1
            ax.fill_between(np.arange(start, end) - 0.6, 0, 1.2, color='k', alpha=0.1)
        for i in range(len(group_counts)):
            start = 0 if i == 0 else group_counts[i - 1]
            end = group_counts[i]
            n = end - start
            text = group_counts.index[i]
            if len(text) > n:
                text = text[:n - 3] + '...'
            ax.annotate(text, (start + n / 2, 1.1), ha='center')
        ax.set_xlim(-1.5, model.k_endog + 0.5)
        ax.annotate('GDP', (model.k_endog - 1.1, 1.05), ha='left', rotation=90)
        fig.tight_layout()
if Dibujar_graficos == 1:
    factor_names = ['Global.1', 'Global.2', 'Confianza', 'Produccion']
    factor_names = rsquared.columns
    mean = vintage_results[Vintage_last].factors.smoothed[factor_names]
    std = pd.concat([vintage_results[Vintage_last].factors.smoothed_cov.loc[name, name] for name in factor_names], axis=1)
    crit = norm.ppf(1 - 0.05 / 2)
    lower = mean - crit * std
    upper = mean + crit * std
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 3))
        mean.plot(ax=ax)
        for name in factor_names:
            ax.fill_between(mean.index, lower[name], upper[name], alpha=0.5)
        ax.set(title='Estimated factors: smoothed estimates and 95% confidence intervals')
        fig.tight_layout()
if Dibujar_graficos == 1:
    prediction_results = vintage_results[Vintage_last].get_prediction(start='2000', end='2030', information_set='predicted')
    variables = ['PMI. MANUFACTURAS. ZONA EURO', 'PMI. SERVICIOS. ZONA EURO', 'PMI. COMPUESTO, TOTAL ECONOMIA ZONA EURO']
    variables = ['GDP q/q (SAAR)']
    point_predictions_GDP = prediction_results.predicted_mean[variables]
    ci = prediction_results.conf_int(alpha=0.25)
    lower_75 = ci[[f'lower {name}' for name in variables]]
    upper_75 = ci[[f'upper {name}' for name in variables]]
    ci = prediction_results.conf_int(alpha=0.5)
    lower = ci[[f'lower {name}' for name in variables]]
    upper = ci[[f'upper {name}' for name in variables]]
    proyeccion_lp = [point_predictions_GDP, lower, upper]
    proyeccion_lp = pd.concat(objs=(iDF for iDF in (point_predictions_GDP, lower_75, upper_75, lower, upper)), axis=1, join='inner').reset_index()
    Book.sheets('Forecast largo plazo').range('b2').options(index=True).value = proyeccion_lp
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 4))
        point_predictions_GDP.loc['2000':datetime.today()].plot(ax=ax)
        point_predictions_GDP.loc[datetime.today():].plot(ax=ax, linestyle='--', color=['C0', 'C1', 'C2'], legend=False)
        forecast_index = point_predictions_GDP.loc[datetime.today():].index
        for name in variables:
            ax.fill_between(forecast_index, lower.loc[forecast_index, f'lower {name}'], upper.loc[forecast_index, f'upper {name}'], alpha=0.55)
            ax.fill_between(forecast_index, lower_75.loc[forecast_index, f'lower {name}'], upper_75.loc[forecast_index, f'upper {name}'], alpha=0.15)
        ylim = ax.get_ylim()
        ylim = ax.set_ylim([-0.05, 0.05])
        ax.vlines(datetime.today(), ylim[0], ylim[1], linewidth=1)
        ax.annotate(' Forecast $\\rightarrow$', ('2020-01', -1.7))
        ax.set(title='GDP US: in-sample predictions and out-of-sample forecasts, with 75% confidence intervals', ylim=ylim)
        fig.tight_layout()
    soft_indicator = 'ISM MANUF'
    gdp_description = 'GDP q/q (SAAR)'
    fcast_m = vintage_results[Vintage_last].forecast('2025-12')[soft_indicator]
    fcast_q = vintage_results[Vintage_last].forecast('2025-12')[gdp_description].resample('Q').last()
    plot_m = pd.concat([dta[Vintage_last].dta_m.loc['2000':, soft_indicator], fcast_m])
    plot_q = pd.concat([dta[Vintage_last].dta_q.loc['2000':, gdp_description], fcast_q])
    with sns.color_palette('deep'):
        fig, axes = plt.subplots(2, figsize=(14, 4))
        plot_q.plot(ax=axes[0])
        axes[0].set(title=gdp_description)
        axes[0].hlines(0, plot_q.index[0], plot_q.index[-1], linewidth=1)
        plot_m.plot(ax=axes[1])
        axes[1].set(title=soft_indicator)
        axes[1].hlines(50, plot_m.index[0], plot_m.index[-1], linewidth=1)
        for i in range(2):
            ylim = axes[i].get_ylim()
            axes[i].fill_between(plot_q.loc[datetime.today():].index, ylim[0], ylim[1], alpha=0.1, color='C0')
            axes[i].annotate(' Forecast $\\rightarrow$', ('2020-03', ylim[0] + 0.1 * ylim[1]))
            axes[i].set_ylim(ylim)
        fig.suptitle('Data and forecasts (February 2020 vintage), transformed scale', fontsize=14, fontweight=600)
        fig.tight_layout(rect=[0, 0, 1, 0.95])
if Dibujar_graficos == 1:
    news_results = {}
    forecast_v = {}
    vintages = Vintages
    group_counts = groups[['description', 'group']]
    group_counts = group_counts[group_counts['description'].isin(dta[Vintage_0].dta_m.columns)]
    group_counts = group_counts.groupby('group', sort=False).count()['description'].cumsum()
    for i in range(1, len(vintages)):
        vintage = vintages[i]
        prev_vintage = vintages[i - 1]
        news_results[vintage] = vintage_results[vintage].news(vintage_results[prev_vintage], impact_date=Fecha_impacto2, impacted_variable=Variable_impactada, comparison_type='previous')
        forecast_v[vintage] = vintage_results[vintage].get_prediction(start='2000', end='2027', information_set='smoothed').predicted_mean[gdp_description].loc[Fecha_impacto2]
    forecast_v = list(forecast_v.values())
    group_impacts = {Vintage_0: None}
    for vintage, news in news_results.items():
        details_by_impact = news.details_by_impact.reset_index().drop(['impact date', 'impacted variable'], axis=1)
        impacts = pd.merge(details_by_impact, groups, how='left', left_on='updated variable', right_on='description').drop('description', axis=1).set_index(['update date', 'updated variable'])
        group_impacts[vintage] = impacts.groupby('group').sum()['impact']
    group_impacts = pd.concat(group_impacts, axis=1).fillna(0).reindex(group_counts.index).T
    group_impacts.index = vintages[1:]
    group_impacts.T.append(group_impacts.sum(axis=1).rename('Total impact on' + Fecha_impacto2 + ' 2024Q2 forecast')).round(2).iloc[:, 1:]
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 6))
        group_impacts.plot(kind='bar', stacked=True, width=0.3, zorder=2, ax=ax)
        x = np.arange(len(group_impacts.index))
        ax.plot(x, forecast_v, marker='o', color='k', markersize=7, linewidth=2)
        ax.hlines(0, -1, len(group_impacts) + 1, linewidth=1)
        labels = group_impacts.index
        ax.xaxis.set_ticklabels(labels)
        ax.xaxis.set_tick_params(size=0)
        ax.xaxis.set_tick_params(labelrotation='auto', labelsize=13)
        ax.yaxis.set_tick_params(direction='in', size=0, labelsize=13)
        ax.yaxis.grid(zorder=0)
        ax.set_title('Evolution of real GDP growth nowcast:' + Fecha_impacto2, fontsize=16, fontweight=600, loc='left')
        [ax.spines[spine].set_visible(False) for spine in ['top', 'left', 'bottom', 'right']]
        ylim = ax.get_ylim()
        ax.vlines(0.5, ylim[0], ylim[1] + 0.005, linestyles='--')
        ax.annotate('Base forecast', (-0.2, 22), fontsize=14)
        ax.annotate('Updated forecasts and impacts from the "news" $\\rightarrow$', (0.65, 22), fontsize=14)
        ax.legend(loc='upper center', ncol=4, fontsize=13, bbox_to_anchor=(0.5, -0.1), frameon=False)
        fig.tight_layout()
'\n####################\n### SACAR FACTORES ###\n####################\n'
print('sacar factores')
factors = {row['description']: [row['group']] for ix, row in groups.iterrows()}
print(factors['ISM MANUF'])
start = '1996'
endog_m_factors = dta[Vintage_last].dta_m.loc[start:, :]
endog_q_factors = dta[Vintage_last].dta_q.loc[start:, :]
"\nendog_m_factors = endog_m_factors.drop(['INDICADOR IFO DEL CLIMA ECONOMICO. ALEMANIA (DATOS CVE)',\n                                        'INDICADOR DE CONFIANZA EN LA INDUSTRIA MANUFACTURERA ZONA EURO (CVE)',\n                                        'INDICADOR DE CONFIANZA DEL COMERCIO MINORISTA ZONA EURO (CVE)',\n                                        'INDICADOR DE CONFIANZA DE LOS CONSUMIDORES ZONA EURO (CVE)',\n                                        'INDICADOR DE CONFIANZASERVICIOS ZONA EURO (CVE)',\n                                        'INDICE PMI INDUSTRIA MANUFACTURERA (CVE). ZONA EURO',\n                                        'INDICE PMI ACTIVIDAD SERVICIOS (CVE). ZONA EURO',\n                                        'INDICE PMI COMPUESTO, TOTAL ECONOMIA (CVE) ZONA EURO',\n                                        'PENDIENTE 10-1'], axis=1)\n"
model_factors = sm.tsa.DynamicFactorMQ(endog_m_factors, endog_quarterly=endog_q_factors, factors=factors, factor_orders=1, factor_multiplicities=1)
model_factors.summary()
results_factors = model_factors.fit(disp=10)
print(results_factors.summary(display_diagnostics=True))
mean = vintage_results[Vintage_last].factors.smoothed
Book.sheets('Factores').range('b5').options(index=True).value = mean
elapsed_time = time.time() - start_time
print(f'Tiempo de ejecución : {elapsed_time:.2f} segundos')