# -*- coding: utf-8 -*-
"""
Created on Thu Jul 10 11:29:05 2025

@author: U61498N
"""

"""
Spyder Editor

This is a temporary script file.

http://www.chadfulton.com/topics/statespace_large_dynamic_factor_models.html

https://www.statsmodels.org/stable/generated/statsmodels.tsa.statespace.dynamic_factor_mq.DynamicFactorMQ.html
https://www.statsmodels.org/devel/examples/notebooks/generated/statespace_news.html

https://www.sr-sv.com/nowcasting-for-financial-markets/
"""

#%matplotlib inline

import types
from datetime import datetime
import numpy as np
import pandas as pd
import statsmodels.api as sm

import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw
import time 

start_time = time.time()


### DATA TRANSFORMATION ###

def transform(column, transforms):
    transformation = transforms[column.name]
    # For quarterly data like GDP, we will compute
    # annualized percent changes
    mult = 1
    #mult = 4 if column.index.freqstr[0] == 'Q' else 1
    
    # 1 => No transformation
    if transformation == 1:
        pass
    # 2 => First difference
    elif transformation == 2:
        column = column.diff()
    # 3 => Second difference
    elif transformation == 3:
        column = column.diff().diff()
    # 4 => Log
    elif transformation == 4:
        column = np.log(column)
    # 5 => Log first difference, multiplied by 100
    #      (i.e. approximate percent change)
    #      with optional multiplier for annualization
    elif transformation == 5:
        column = np.log(column).diff() * 100 * mult
    # 6 => Log second difference, multiplied by 100
    #      with optional multiplier for annualization
    elif transformation == 6:
        column = np.log(column).diff().diff() * 100 * mult
    # 7 => Exact percent change, multiplied by 100
    #      with optional annualization
    elif transformation == 7:
        column = ((column / column.shift(1))**mult - 1.0) * 100
    # 8 => Variacion anual de indicador mensual
    elif transformation == 8:
        column = ((column / column.shift(12))**mult - 1.0) * 100        
    # 11 => LAGS 1
    elif transformation == 11:
        column = column.shift(1)
    # 12 => LAGS 2
    elif transformation == 12:
        column = column.shift(2)
    # 13 => LAGS 3
    elif transformation == 13:
        column = column.shift(3)
    # 14 => LAGS 4
    elif transformation == 14:
        column = column.shift(4)
    # 15 => LAGS 5
    elif transformation == 15:
        column = column.shift(5)
    # 112 => LAGS 12
    elif transformation == 112:
        column = column.shift(12)
    return column


### REMOVE OUTLIERS ###

def remove_outliers(dta):
    # Compute the mean and interquartile range
    mean = dta.mean()
    iqr = dta.quantile([0.25, 0.75]).diff().T.iloc[:, 1]
    #iqr = dta.quantile([0.35, 0.65]).diff().T.iloc[:, 1]
    # Replace entries that are more than 10 times the IQR
    # away from the mean with NaN (denotes a missing entry)
    mask = np.abs(dta) > mean + 10 * iqr
    treated = dta.copy()
    treated[mask] = np.nan

    return treated
"""
####################
### LOADING DATA ###
####################
"""
#path = 'R:/_BankAnalytics/Centros/CoyEcono/Datos/publico/Modelizacion/Factor model/Eurozone/'
#path = 'C:/Users/u61498n/Desktop/Factor model/Eurozone/'
path = 'Y:/mesa teso/JRS/8_DFM Nowcasting_US/'
path += 'Datos_US.xlsx'
Book = xw.Book(path)

# Opciones
Dibujar_graficos = Book.sheets('Indice').range('c5').options(index=False).value
#Trimestre_objetivo = Book.sheets('Indice').range('c7').options(index=False).value
#Trimestre_objetivo_text = "2022Q4"

#Vintages =

Vintages = Book.sheets('Indice').range('c6').options(expand='right').value
""" 
fecha_hoy =  datetime.today().strftime('%d-%m-%Y') # Nueva hoja con la vintage de hoy
Vintages = Vintages+[fecha_hoy]
Book.sheets('Indice').range('c6').value = Vintages #añado la vintage de hoy a la lista de vintages
"""
Vintage_0 = Vintages[0] # Estimaremos con esta Vintage y actualizaremos con los datos entrados en las siguientes vintages
Vintage_last=Vintages[-1]

# leo los datos vinculados a la descarga con freq mensual
Datos_vinculados = Book.sheets('Datos').range('A1')\
        .options(pd.DataFrame,expand='table',index=False,decimal='.').value
# y los pego como valor en la hoja de la vintage a estimar
try:
    Book.sheets.add(Vintage_last)
except:
  print("Hoja ya creada")

Book.sheets(Vintage_last).range('a1').options(index=False).value = Datos_vinculados


### Datos Mensuales

# 1. Lectura de datos

def load_fredmd_data(vintage):
    orig_m = Book.sheets(vintage).range('A1')\
            .options(pd.DataFrame,expand='table',index=False,decimal='.').value
    orig_m=orig_m.drop(columns=['SERIES'])
    orig_m=orig_m.drop(columns=['GDP level'])
    
    
    # 2. Extraer informacion de transformacion y de los factores
    transform_m = orig_m.iloc[1, 1:]
    factor_m1 = orig_m.iloc[2, 1:]
    factor_m2 =orig_m.iloc[3, 1:]
    orig_m = orig_m.iloc[4:]
    
    # 3. Extract the date as an index
    #orig_m = orig_m.set_index(pd.DatetimeIndex(orig_m['fecha']))
    orig_m.index = pd.PeriodIndex(orig_m.fecha.tolist(), freq='M')
    orig_m = orig_m.iloc[:,1:]
    
    # 4. Apply the transformations
    for i in orig_m.columns:
        orig_m[i] = pd.to_numeric(orig_m[i],errors = 'coerce')
        
    dta_m = orig_m.apply(transform, axis=0,
                           transforms=transform_m)
    
    # 5. Remove outliers (but not in 2020)
    #dta_m.loc[:] = remove_outliers(dta_m.loc[:])
    dta_m.loc[:'2019-12'] = remove_outliers(dta_m.loc[:'2019-12'])

    
    ### Datos Trimestrales
    
    # 1. Lectura de datos
    orig_q = Book.sheets(vintage).range('A1')\
            .options(pd.DataFrame,expand='table',index=False,decimal='.').value
    orig_q=orig_q[['fecha','GDP level']]
    orig_q=orig_q.dropna()
    
    # 2. Extraer informacion de transformacion
    transform_q = orig_q.iloc[1, 1:]
    factor_q1 = orig_q.iloc[2, 1:]
    factor_q2 =orig_q.iloc[3, 1:]
    orig_q = orig_q.iloc[4:]
    
    # 3. Extract the date as an index
    #orig_m = orig_m.set_index(pd.DatetimeIndex(orig_m['fecha']))
    orig_q.index = pd.PeriodIndex(orig_q.fecha.tolist(), freq='Q')
    orig_q = orig_q.iloc[:,1:]
    
    # 4. Apply the transformations
    for i in orig_q.columns:
        orig_q[i] = pd.to_numeric(orig_q[i],errors = 'coerce')
        
    dta_q = orig_q.apply(transform, axis=0,
                           transforms=transform_q)
    
    # 5. Remove outliers (but not in 2020)
    #dta_q.loc[:] = remove_outliers(dta_q.loc[:])
    dta_q.loc[:'2019-12'] = remove_outliers(dta_q.loc[:'2019-12'])
    
     # - Output datasets ------------------------------------------------------
    return types.SimpleNamespace(
        orig_m=orig_m, orig_q=orig_q,
        dta_m=dta_m, transform_m=transform_m,factor_m2=factor_m2,
        dta_q=dta_q, transform_q=transform_q, factor_q2=factor_q2)

dta = {date: load_fredmd_data(date)
       for date in Vintages}

#pegamos muestra transformada en cada vintaje para comparar con forecast
endog_qmonthly=dta[Vintage_last].dta_q
endog_qmonthly = endog_qmonthly.asfreq('M')
Muestra_transformada = dta[Vintage_last].dta_m.join(endog_qmonthly, lsuffix='_')
Book.sheets(Vintage_last).range('bc5').options(index=False).value = Muestra_transformada
#for vintage in Vintages:
#    Book.sheets(Vintage_last).range('bc5').options(index=False).value = dta[vintage].value


### Grupos
groups = dta[Vintage_0].factor_m2.copy()
groups = groups.append(dta[Vintage_0].factor_q2)
groups=pd.DataFrame(groups).reset_index()
groups.columns = ['description', 'group']

# Dibujar numero de grupos
if Dibujar_graficos==1:
    # Display the number of variables in each group
    (groups.groupby('group', sort=False)
           .count()
           .rename({'description': '# series in group'}, axis=1))
    
    # Construct the variable => list of factors dictionary
    factors = {row['description']: ['Global', row['group']]
               for ix, row in groups.iterrows()}
    #factors = {row['description']: [row['group']]
    #           for ix, row in groups.iterrows()}
    # Check that we have the desired factor 
    print(factors['ISM MANUF'])

# Factor multiplicities
factor_multiplicities = {'Global': 1}

# Factor orders

factor_orders = {
    ('Hard', 'Soft'): 1,
    'Global': 2}
#factor_orders = {('Hard', 'Soft'): 1}


"""
####################
###    MODELO    ###
####################
"""
# Get the baseline monthly and quarterly datasets
start = '1996'
endog_m = dta[Vintage_0].dta_m.loc[start:, :]
endog_q = dta[Vintage_0].dta_q.loc[start:, :]


# Construct the dynamic factor model
"""
model = sm.tsa.DynamicFactorMQ(
    endog_m, endog_quarterly=endog_q,
    factors=factors, factor_orders=factor_orders,
    factor_multiplicities=factor_multiplicities)
"""
model = sm.tsa.DynamicFactorMQ(
    endog_m, endog_quarterly=endog_q,
    factors=1, factor_orders=1,
    factor_multiplicities=1)


model.summary()
results = model.fit(disp=10)
print(results.summary(display_diagnostics=True))


"""
####################
### Forecasting ###
####################
"""
print("FORECASTING en cada Vintage")
prediction_results = results.get_prediction(start='2000', end='2027')
point_predictions = prediction_results.predicted_mean
Book.sheets(Vintage_0).range('aa65').options(index=True).value = point_predictions

### Actualización con los nuevos datos de las sucesivas Vintages
vintage_results = {Vintage_0: results}
for vintage in Vintages:
    print (vintage)
    # Get updated data for the vintage
    updated_endog_m = dta[vintage].dta_m.loc[start:, :]
    updated_endog_q = dta[vintage].dta_q.loc[start:, :]
    # Get updated results for for the vintage
    vintage_results[vintage] = results.apply(
        updated_endog_m, endog_quarterly=updated_endog_q)
    prediction_results = vintage_results[vintage].get_prediction(start='2000', end='2027')
    point_predictions = prediction_results.predicted_mean
    Book.sheets(vintage).range('aa65').options(index=True).value = point_predictions



"""
####################
####### NEWS #######
####################
"""

print("NEWS en cada Vintage y para 3 fechas especificas: Q+1, Q+2 y Q+3")

Variable_impactada = 'GDP level'
Fecha_impacto1 = Book.sheets('Impactos').range('d3').options(index=False).value
Fecha_impacto2 = Book.sheets('Impactos').range('d47').options(index=False).value
Fecha_impacto3 = Book.sheets('Impactos').range('d87').options(index=False).value

# Lista de fechas de impacto
fechas_impacto = [Fecha_impacto1, Fecha_impacto2, Fecha_impacto3]  # Añade todas las fechas que necesites

# Lista de celdas de inicio para cada fecha de impacto
celdas_inicio = ['cb5', 'co5', 'db5']

# Lista de rangos de celdas a limpiar para cada fecha de impacto
rangos_limpiar = ['cb5:ck25', 'co5:cx25', 'db5:dek25']

for fecha, celda, rango in zip(fechas_impacto, celdas_inicio, rangos_limpiar):
    print(fecha)
    for i in range(1, len(Vintages)):
        vintage = Vintages[i]
        prev_vintage = Vintages[i - 1]

        # Imprimir el tamaño del índice del modelo actual
        print(vintage_results[vintage].model._index.size)
        
        # Calcular las noticias entre el vintage actual y el anterior
        news = vintage_results[vintage].news(
            vintage_results[prev_vintage], 
            impact_date=fecha,
            impacted_variable=Variable_impactada,
            comparison_type='previous')
        
        # The `summary` method summarizes all updates. Here we aren't
        # showing it, to save space.
        # news.summary()

        # Obtener los detalles por impacto
        details = news.details_by_impact
        
        # Calcular el impacto absoluto y ordenar por peso
        details['absolute impact'] = np.abs(details['impact'])
        details = details.sort_values('weight', ascending=False)
        
        # Limpiar el rango de celdas específico para cada fecha de impacto
        Book.sheets(vintage).range(rango).clear_contents()
        
        # Escribir los detalles en el libro en la celda especificada
        Book.sheets(vintage).range(celda).options(index=True).value = details
"""

############################################
####### NEWS para yn unico trimestre #######
############################################
Variable_impactada = 'GDP level'
Fecha_impacto1 = Book.sheets('Impactos').range('d3').options(index=False).value
for i in range(1, len(Vintages)):
    vintage = Vintages[i]
    prev_vintage = Vintages[i - 1]

    print(vintage_results[vintage].model._index.size)
    news = vintage_results[vintage].news(
    vintage_results[prev_vintage], impact_date=Fecha_impacto1,
    impacted_variable=Variable_impactada,
    comparison_type='previous')

    # The `summary` method summarizes all updates. Here we aren't
    # showing it, to save space.
    # news.summary()
    
    details = news.details_by_impact
    #details.index = details.index.droplevel(['update date','impact date', 'impacted variable'])
    details['absolute impact'] = np.abs(details['impact'])
    details = (details.sort_values('weight', ascending=False))
    #print(details)
    Book.sheets(vintage).range("ca5:czm25").clear_contents()
    Book.sheets(vintage).range('cb5').options(index=True).value = details
    
"""    
"""
####################
####### NEWS pendiente #######
####################

Variable_impactada = 'PENDIENTE 10-1'
Fecha_impacto = Book.sheets('Impactos_Pendiente').range('d3').options(index=False).value

for i in range(1, len(Vintages)):
    vintage = Vintages[i]
    prev_vintage = Vintages[i - 1]

    print(vintage_results[vintage].model._index.size)
    news = vintage_results[vintage].news(
    vintage_results[prev_vintage], impact_date=Fecha_impacto,
    impacted_variable=Variable_impactada,
    comparison_type='previous')

    # The `summary` method summarizes all updates. Here we aren't
    # showing it, to save space.
    # news.summary()
    
    details = news.details_by_impact
    #details.index = details.index.droplevel(['update date','impact date', 'impacted variable'])
    details['absolute impact'] = np.abs(details['impact'])
    details = (details.sort_values('weight', ascending=False))
    #print(details)
    Book.sheets(vintage).range("cn5:czm25").clear_contents()
    Book.sheets(vintage).range('co5').options(index=True).value = details
    
"""

"""
####################
### POST ESTIMACIÓN ###
####################
"""

# Dibujar capacidad explicativa de los factores
if Dibujar_graficos==1:
    ### Explanatory power of the factors
    method='individual'# retrieves the R2 value for each observed variable regressed on each individual factor (plus a constant term)
    method='joint' # retrieves the R2 value for each observed variable regressed on all factors that the variable loads on
    method='cumulative'# retrieves the R2 value for each observed variable regressed on an expanding set of factors.
    
    rsquared = results.get_coefficients_of_determination(method='individual')
    
    top_ten = []
    for factor_name in rsquared.columns[:1]:
        top_factor = (rsquared[factor_name].sort_values(ascending=False)
                                           .iloc[:10].round(2).reset_index())
        top_factor.columns = pd.MultiIndex.from_product([
            [f'Top ten variables explained by {factor_name}'],
            ['Variable', r'$R^2$']])
        top_ten.append(top_factor)
    pd.concat(top_ten, axis=1)
    
    with sns.color_palette('deep'):
        fig = results.plot_coefficients_of_determination(method='individual', figsize=(14, 9))
        fig.suptitle(r'$R^2$ - regression on individual factors', fontsize=14, fontweight=600)
        fig.tight_layout(rect=[0, 0, 1, 0.95]);
    
    #group_counts = defn_m[['description', 'group']]
    #group_counts = group_counts[group_counts['description'].isin(dta['2020-02'].dta_m.columns)]
    group_counts = groups.groupby('group', sort=False).count()['description'].cumsum()
    
    with sns.color_palette('deep'):
        fig = results.plot_coefficients_of_determination(method='joint', figsize=(14, 3));
    
        # Add in group labels
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
    
        # Add label for GDP
        ax.set_xlim(-1.5, model.k_endog + 0.5)
        ax.annotate('GDP', (model.k_endog - 1.1, 1.05), ha='left', rotation=90)
    
        fig.tight_layout();

# Dibujar factores:
if Dibujar_graficos==1:
    # Get estimates of the global and labor market factors,
    # conditional on the full dataset ("smoothed")
    factor_names = ['Global.1', 'Global.2', 'Confianza','Produccion']
    factor_names = rsquared.columns
    mean = vintage_results[Vintage_last].factors.smoothed[factor_names]
    
    # Compute 95% confidence intervals
    from scipy.stats import norm
    std = pd.concat([vintage_results[Vintage_last].factors.smoothed_cov.loc[name, name]
                     for name in factor_names], axis=1)
    crit = norm.ppf(1 - 0.05 / 2)
    lower = mean - crit * std
    upper = mean + crit * std
    
    
    
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 3))
        mean.plot(ax=ax)
        
        for name in factor_names:
            ax.fill_between(mean.index, lower[name], upper[name], alpha=0.5)
        
        ax.set(title='Estimated factors: smoothed estimates and 95% confidence intervals')
        fig.tight_layout();




# Dibujar ejemplo de forecasting para PMIs
if Dibujar_graficos==1:
    # Note: these forecasts are in the same scale as the variables passed to the DynamicFactorMQ constructor, even if standardize=True has been used.   
    # Create forecasts results objects, through the end of 20201
    prediction_results = vintage_results[Vintage_last].get_prediction(start='2000', end='2030')
    
    variables = ['CONF CONSUM (U. MICH)',
                 'BUSINESS OUTLOOK (PHIL)',
                 'ISM MANUF']
    
    variables = ['GDP level']
    
    # The `predicted_mean` attribute gives the same
    # point forecasts that would have been returned from
    # using the `predict` or `forecast` methods.
    point_predictions_GDP = prediction_results.predicted_mean[variables]
    
    # We can use the `conf_int` method to get confidence
    # intervals; here, the 95% confidence interval
    ci = prediction_results.conf_int(alpha=0.25)
    lower_75 = ci[[f'lower {name}' for name in variables]]
    upper_75 = ci[[f'upper {name}' for name in variables]]
    ci = prediction_results.conf_int(alpha=0.5)
    lower = ci[[f'lower {name}' for name in variables]]
    upper = ci[[f'upper {name}' for name in variables]]
    
    proyeccion_lp=[point_predictions_GDP,lower,upper]
    proyeccion_lp=pd.concat(objs=(iDF for iDF in (point_predictions_GDP,lower_75,upper_75,lower,upper)),axis=1,join='inner').reset_index()
    Book.sheets('Forecast largo plazo').range('b2').options(index=True).value = proyeccion_lp

    
    # Plot the forecasts and confidence intervals
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 4))
    
        # Plot the in-sample predictions
        point_predictions_GDP.loc['2022-01':datetime.today()].plot(ax=ax)
    
        # Plot the out-of-sample forecasts
        point_predictions_GDP.loc[datetime.today():].plot(ax=ax, linestyle='--',
                                               color=['C0', 'C1', 'C2'],
                                               legend=False)
        
        # Confidence intervals
        forecast_index = point_predictions_GDP.loc[datetime.today():].index

        for name in variables:
            ax.fill_between(forecast_index,
                            lower.loc[forecast_index, f'lower {name}'],
                            upper.loc[forecast_index, f'upper {name}'], alpha=0.55)
            ax.fill_between(forecast_index,
                            lower_75.loc[forecast_index, f'lower {name}'],
                            upper_75.loc[forecast_index, f'upper {name}'], alpha=0.15)
            
        # Forecast period, set title
        ylim = ax.get_ylim()
        ylim = ax.set_ylim([-0.01, 0.01])
        ax.vlines(datetime.today(), ylim[0], ylim[1], linewidth=1)
        ax.annotate(r' Forecast $\rightarrow$', ('2020-01', -1.7))
        ax.set(title=('GDP Eurozona:'
                      ' in-sample predictions and out-of-sample forecasts, with 75% confidence intervals'), ylim=ylim)
        
        fig.tight_layout()
    
    # Compute the point forecasts
    soft_indicator = 'ISM MANUF'
    gdp_description = 'GDP level'
    
    
    fcast_m = vintage_results[Vintage_last].forecast('2025-12')[soft_indicator]
    fcast_q = vintage_results[Vintage_last].forecast('2025-12')[gdp_description].resample('Q').last()    
    # For more convenient plotting, combine the observed data with the forecasts
    plot_m = pd.concat([dta[Vintage_last].dta_m.loc['2000':, soft_indicator], fcast_m])
    plot_q = pd.concat([dta[Vintage_last].dta_q.loc['2000':, gdp_description], fcast_q])
    
    with sns.color_palette('deep'):
        fig, axes = plt.subplots(2, figsize=(14, 4))
    
        # Plot real GDP growth, data and forecasts
        plot_q.plot(ax=axes[0])
        axes[0].set(title=gdp_description)
        axes[0].hlines(0, plot_q.index[0], plot_q.index[-1], linewidth=1)
    
        # Plot the change in the unemployment rate, data and forecasts
        plot_m.plot(ax=axes[1])
        axes[1].set(title=soft_indicator)
        axes[1].hlines(50, plot_m.index[0], plot_m.index[-1], linewidth=1)
        
        # Show the forecast period in each graph
        for i in range(2):
            ylim = axes[i].get_ylim()
            axes[i].fill_between(plot_q.loc[datetime.today():].index,
                                 ylim[0], ylim[1], alpha=0.1, color='C0')
            axes[i].annotate(r' Forecast $\rightarrow$',
                             ('2020-03', ylim[0] + 0.1 * ylim[1]))
            axes[i].set_ylim(ylim)
    
        # Title
        fig.suptitle('Data and forecasts (February 2020 vintage), transformed scale',
                     fontsize=14, fontweight=600)
    
        fig.tight_layout(rect=[0, 0, 1, 0.95]);
        
        

# Dibujar contribuciones a news por grupos
if Dibujar_graficos==1:
    
    news_results = {}
    vintages = Vintages
    
    group_counts = groups[['description', 'group']]
    group_counts = group_counts[group_counts['description'].isin(dta[Vintage_0].dta_m.columns)]
    group_counts = group_counts.groupby('group', sort=False).count()['description'].cumsum()
    
    
    
    for i in range(1, len(vintages)):
        vintage = vintages[i]
        prev_vintage = vintages[i - 1]
    
        # Notice that to get the "incremental" news, we are computing
        # the news relative to the previous vintage and not to the baseline
        # (February 2020) vintage
        news_results[vintage] = vintage_results[vintage].news(
            vintage_results[prev_vintage],
            impact_date=Fecha_impacto1,
            impacted_variable=Variable_impactada,
            comparison_type='previous')
    
    
    
    group_impacts = {Vintage_0: None}
    for vintage, news in news_results.items():
        # Start from the details by impact table
        details_by_impact = (
            news.details_by_impact.reset_index()
                .drop(['impact date', 'impacted variable'], axis=1))
        
        # Merge with the groups dataset, so that we can identify
        # which group each individual impact belongs to
        impacts = (pd.merge(details_by_impact, groups, how='left',
                            left_on='updated variable', right_on='description')
                     .drop('description', axis=1)
                     .set_index(['update date', 'updated variable']))
    
        # Compute impacts by group, summing across the individual impacts
        group_impacts[vintage] = impacts.groupby('group').sum()['impact']
    
    group_impacts = (
        pd.concat(group_impacts, axis=1)
          .fillna(0)
          .reindex(group_counts.index).T)
    group_impacts.index = vintages[1:]
    
    (group_impacts.T
        .append(group_impacts.sum(axis=1).rename('Total impact on'+Fecha_impacto1+' 2024Q2 forecast'))
        .round(2).iloc[:, 1:])
    
    
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 6))
    
        # Stacked bar plot showing the impacts by group
        group_impacts.plot(kind='bar', stacked=True, width=0.3, zorder=2, ax=ax);
    
        # Line plot showing the forecast for real GDP growth in 2020Q2 for each vintage
        x = np.arange(len(group_impacts.index))
        ax.plot(x, point_predictions.loc[group_impacts.index][gdp_description], marker='o', color='k', markersize=7, linewidth=2)
        ax.hlines(0, -1, len(group_impacts) + 1, linewidth=1)
    
        # x-ticks
        labels = group_impacts.index#.strftime('%b')
        ax.xaxis.set_ticklabels(labels)
        ax.xaxis.set_tick_params(size=0)
        ax.xaxis.set_tick_params(labelrotation='auto', labelsize=13)
    
        # y-ticks
        ax.yaxis.set_tick_params(direction='in', size=0, labelsize=13)
        ax.yaxis.grid(zorder=0)
        
        # title, remove spines
        ax.set_title('Evolution of real GDP growth nowcast:'+Fecha_impacto1, fontsize=16, fontweight=600, loc='left')
        [ax.spines[spine].set_visible(False)
         for spine in ['top', 'left', 'bottom', 'right']]
        
        # base forecast vs updates
        ylim = ax.get_ylim()
        ax.vlines(0.5, ylim[0], ylim[1] + 0.005, linestyles='--')
        ax.annotate('Base forecast', (-0.2, 22), fontsize=14)
        ax.annotate(r'Updated forecasts and impacts from the "news" $\rightarrow$', (0.65, 22), fontsize=14)
    
        # legend
        ax.legend(loc='upper center', ncol=4, fontsize=13, bbox_to_anchor=(0.5, -0.1), frameon=False)
    
        fig.tight_layout();
        
"""
####################
### SACAR FACTORES ###
####################
"""        
print('sacar factores')
# Construct the variable => list of factors dictionary

factors = {row['description']: [row['group']]
           for ix, row in groups.iterrows()}
# Check that we have the desired factor 
print(factors['ISM MANUF'])  
start = '1996'
#endog_m = dta[Vintage_0].dta_m.loc[start:, :]
#endog_q = dta[Vintage_0].dta_q.loc[start:, :]

endog_m_factors = dta[Vintage_last].dta_m.loc[start:, :]
endog_q_factors = dta[Vintage_last].dta_q.loc[start:, :]
#endog_m_factors = endog_m_factors.drop(['TASA DE PARO ZONA EURO (CVE)','IPI ZONA EURO, EXCLUÍDA CONSTRUCCIÓN (CVE)','INDICE BURSATIL EUROSTOXX 50 (BASE 31/12/1991 = 1000)'], axis=1)
"""
endog_m_factors = endog_m_factors.drop(['INDICADOR IFO DEL CLIMA ECONOMICO. ALEMANIA (DATOS CVE)',
                                        'INDICADOR DE CONFIANZA EN LA INDUSTRIA MANUFACTURERA ZONA EURO (CVE)',
                                        'INDICADOR DE CONFIANZA DEL COMERCIO MINORISTA ZONA EURO (CVE)',
                                        'INDICADOR DE CONFIANZA DE LOS CONSUMIDORES ZONA EURO (CVE)',
                                        'INDICADOR DE CONFIANZASERVICIOS ZONA EURO (CVE)',
                                        'INDICE PMI INDUSTRIA MANUFACTURERA (CVE). ZONA EURO',
                                        'INDICE PMI ACTIVIDAD SERVICIOS (CVE). ZONA EURO',
                                        'INDICE PMI COMPUESTO, TOTAL ECONOMIA (CVE) ZONA EURO',
                                        'PENDIENTE 10-1'], axis=1)
"""
model_factors = sm.tsa.DynamicFactorMQ(
    endog_m_factors, endog_quarterly=endog_q_factors,
    factors=factors, factor_orders=1,
    factor_multiplicities=1)


model_factors.summary()
results_factors = model_factors.fit(disp=10)
print(results_factors.summary(display_diagnostics=True))
      
#updated_endog_m = dta[vintage].dta_m.loc[start:, :]
#updated_endog_q = dta[vintage].dta_q.loc[start:, :]
##### Get updated results for for the vintage
#vintage_results[vintage] = results.apply(updated_endog_m, endog_quarterly=updated_endog_q) 

#mean = vintage_results[Vintage_last].factors.smoothed#['Hard','Soft']
mean = results_factors.factors.smoothed#['Hard','Soft']
Book.sheets("Factores").range('b5').options(index=True).value = mean       

"""
# sacar forecasting_0 
point_forecasts = results.forecast(steps=24)
print(point_forecasts)
prediction_results = results.get_prediction(start='2000', end='2023')
point_predictions = prediction_results.predicted_mean

#♦ intervalos
ci = prediction_results.conf_int(alpha=0.50)
lower = ci[[f'lower {name}' for name in Muestra_transformada.columns]]
upper = ci[[f'upper {name}' for name in Muestra_transformada.columns]]

pronosticos = point_predictions.join(lower, lsuffix='_').join(upper, lsuffix='_')
Book.sheets(Vintage_0).range('dc5').options(index=False).value = pronosticos


### Focus on GDP: Updating vintages
# The original point forecasts are monthly
point_forecasts_m = results.forecast(Trimestre_objetivo)[gdp_description]

# Resample to quarterly frequency by taking the value in the last
# month of each quarter
point_forecasts_q = point_forecasts_m.resample('Q').last()

print('Baseline', Vintage_0, 'forecast for real GDP growth in', Trimestre_objetivo_text,'is',point_forecasts_q[Trimestre_objetivo_text])


# Since we will be collecting results for a number of vintages,
# construct a dictionary to hold them, and include the baseline
# results from Vintage_0
vintage_results = {Vintage_0: results}

# Get the updated monthly and quarterly datasets
start = '2000'
updated_endog_m = dta[Vintage_last].dta_m.loc[start:, :]
#gdp_description = defn_q.loc['GDPC1', 'description']
updated_endog_q = dta[Vintage_last].dta_q.loc[start:, :]

# Get the results para la ultima actualizacion de datos using `apply`
vintage_results[Vintage_last] = results.apply(
    updated_endog_m, endog_quarterly=updated_endog_q)

# Print the updated forecast for real GDP growth in 2020Q2
updated_forecasts_q = vintage_results[Vintage_last].get_prediction(Trimestre_objetivo).predicted_mean[gdp_description].resample('Q').last()

print('Ultima actualización a ',Vintage_last,' forecast for real GDP growth in',Trimestre_objetivo_text,'is',updated_forecasts_q[Trimestre_objetivo_text])

# Apply our results to the remaining vintages
for vintage in Vintages:
    print (vintage)
    # Get updated data for the vintage
    updated_endog_m = dta[vintage].dta_m.loc[start:, :]
    updated_endog_q = dta[vintage].dta_q.loc[start:, :]
    # Get updated results for for the vintage
    vintage_results[vintage] = results.apply(
        updated_endog_m, endog_quarterly=updated_endog_q)


# Compute forecasts for each vintage
forecasts = {vintage: res.get_prediction(Trimestre_objetivo).predicted_mean[gdp_description]
                         .resample('Q').last().loc[Trimestre_objetivo_text]
             for vintage, res in vintage_results.items()}
# Convert to a Pandas series with a date index
forecasts = pd.Series(list(forecasts.values()), index=pd.PeriodIndex(forecasts.keys(), freq='M'))
    
# Print our forecast for 2020Q2 real GDP growth across all vintages
for vintage, value in forecasts.items():
    print(f'{vintage} forecast for real GDP growth in',Trimestre_objetivo_text,' 2020Q2:'
          f' {value:.2f}%')

# Computar impactos

# Compute the news and impacts on the real GDP growth forecast
# for 2020Q2, between the April and March vintages

for i in range(1, len(Vintages)):
    vintage = Vintages[i]
    prev_vintage = Vintages[i - 1]

    news = vintage_results[vintage].news(
    vintage_results[prev_vintage], impact_date='2022-09',
    impacted_variable=Variable_impactada,
    comparison_type='previous')

    # The `summary` method summarizes all updates. Here we aren't
    # showing it, to save space.
    # news.summary()
    
    details = news.details_by_impact
    details.index = details.index.droplevel(['update date','impact date', 'impacted variable'])
    details['absolute impact'] = np.abs(details['impact'])
    details = (details.sort_values('weight', ascending=False))
    print(details)


    Book.sheets(vintage).range('aa5').options(index=True).value = details
    
impactos=()
for i in Vintages:
    impactos_vintage=Book.sheets(i).range('ai5').options(index=True,expand='table').value

    impactos = (pd.merge(impactos, impactos_vintage, how='left'))#,
                        left_on='updated variable', right_on='description')
                 .drop('description', axis=1)
                 .set_index(['update date', 'updated variable']))    
    
Book.sheets('Datos').range('A1')\
        .options(pd.DataFrame,expand='table',index=False,decimal='.').value    
    
news_results = {}
vintages=Vintages
impact_date = Trimestre_objetivo_text

for i in range(1, len(vintages)):
    vintage = vintages[i]
    prev_vintage = vintages[i - 1]

    # Notice that to get the "incremental" news, we are computing
    # the news relative to the previous vintage and not to the baseline
    # (February 2020) vintage
    news_results[vintage] = vintage_results[vintage].news(
        vintage_results[prev_vintage],
        impact_date=impact_date,
        impacted_variable=gdp_description,
        comparison_type='previous')

group_impacts = {Vintage_0: None}    
for vintage, news in news_results.items():
    print(vintage)
    # Start from the details by impact table
    details_by_impact = (
        news.details_by_impact.reset_index()
            .drop(['impact date', 'impacted variable'], axis=1))
    
    # Merge with the groups dataset, so that we can identify
    # which group each individual impact belongs to
    impacts = (pd.merge(details_by_impact, groups, how='left',
                        left_on='updated variable', right_on='description')
                 .drop('description', axis=1)
                 .set_index(['update date', 'updated variable']))
    # Compute impacts by group, summing across the individual impacts
    group_impacts[vintage] = impacts.groupby('updated variable').sum()['impact']

# Add in a row of zeros for the baseline forecast
group_impacts[Vintage_0] = group_impacts[vintages[1]] * np.nan

# Convert into a Pandas DataFrame, and fill in missing entries
# with zeros (missing entries happen when there were no updates
# for a given group in a given vintage)
group_impacts = (
    pd.concat(group_impacts, axis=1)
      .fillna(0)
      .reindex(group_counts.index).T)
group_impacts.index = forecasts.index

# Print the table of impacts from data in each group,
# along with a row with the "Total" impact
(group_impacts.T
    .append(group_impacts.sum(axis=1).rename('Total impact on 2020Q2 forecast'))
    .round(2).iloc[:, 1:])




Book.sheets('Grafico').range('q200').options(index=True).value = impacts.round(2)





pred = vintage_results[Vintages[-1]].get_prediction(start='2000', end='2023')
pred_mean = pred.predicted_mean[gdp_description].resample('Q').last()
    
pred_ci = pred.conf_int()[['lower PIB PRECIOS CONSTANTES ZONA EURO', 'upper PIB PRECIOS CONSTANTES ZONA EURO']].resample('Q').last()
pred_ci = pred_ci['2021Q4':]    

plt.figure()
ax = endog_q['2015Q1':].plot(label='Observado',figsize=(15, 15))
pred_mean['2021Q4':].plot(ax=ax, label='Dynamic forecast')

ax.fill_between(pred_ci.index,
                pred_ci.iloc[:, 0],
                pred_ci.iloc[:, 1], color='k', alpha=.25)
ax.fill_betweenx(ax.get_ylim(), pd.to_datetime(['2021Q4']), endog_q.index[-1],
                 alpha=.1, zorder=-1)
ax.set_title([gdp_description])
ax.set_xlabel('Date')
ax.set_ylabel('PIB q/q')
plt.legend()

plt.show()





if Dibujar_graficos==1:
    # Note: these forecasts are in the same scale as the variables passed to the DynamicFactorMQ constructor, even if standardize=True has been used.   
    # Create forecasts results objects, through the end of 20201
    prediction_results = vintage_results[Vintages[-1]].get_prediction(start='2000', end='2023')
    
    variables = [gdp_description]
    
    # The `predicted_mean` attribute gives the same
    # point forecasts that would have been returned from
    # using the `predict` or `forecast` methods.
    point_predictions = prediction_results.predicted_mean[variables].resample('Q').last()
    
    # We can use the `conf_int` method to get confidence
    # intervals; here, the 95% confidence interval
    ci = prediction_results.conf_int(alpha=0.05).resample('Q').last()
    lower = ci[[f'lower {name}' for name in variables]].resample('Q').last()
    upper = ci[[f'upper {name}' for name in variables]].resample('Q').last()
    
    # Plot the forecasts and confidence intervals
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 4))
    
        # Plot the in-sample predictions
        point_predictions.loc['2021Q1':'2021Q3'].plot(ax=ax)
    
        # Plot the out-of-sample forecasts
        point_predictions.loc['2021Q3':].plot(ax=ax, linestyle='--',
                                               color=['C0', 'C1', 'C2'],
                                               legend=False)
    
        # Confidence intervals
        for name in variables:
            ax.fill_between(ci.index.loc['2021Q3':],
                            lower[f'lower {name}'],
                            upper[f'upper {name}'], alpha=0.1)
            
        # Forecast period, set title
        ylim = ax.get_ylim()
        ax.vlines('2021Q3', ylim[0], ylim[1], linewidth=1)
        ax.annotate(r' Forecast $\rightarrow$', ('2020Q1', -1.7))
        ax.set(title=('PMIs Eurozona:'
                      ' in-sample predictions and out-of-sample forecasts, with 95% confidence intervals'), ylim=ylim)
        
        fig.tight_layout()


news_results = {}
vintages = Vintages
impact_date = '2021-12'

for i in range(1, len(vintages)):
    vintage = vintages[i]
    prev_vintage = vintages[i - 1]

    # Notice that to get the "incremental" news, we are computing
    # the news relative to the previous vintage and not to the baseline
    # (February 2020) vintage
    news_results[vintage] = vintage_results[vintage].news(
        vintage_results[prev_vintage],
        impact_date=impact_date,
        impacted_variable=gdp_description,
        comparison_type='previous')

group_impacts = {Vintage_0: None}

for vintage, news in news_results.items():
    # Start from the details by impact table
    details_by_impact = (
        news.details_by_impact.reset_index()
            .drop(['impact date', 'impacted variable'], axis=1))
    
    # Merge with the groups dataset, so that we can identify
    # which group each individual impact belongs to
    impacts = (pd.merge(details_by_impact, groups, how='left',
                        left_on='updated variable', right_on='description')
                 .drop('description', axis=1)
                 .set_index(['update date', 'updated variable']))

    # Compute impacts by group, summing across the individual impacts
    group_impacts[vintage] = impacts.groupby('group').sum()['impact']

# Add in a row of zeros for the baseline forecast
group_impacts[Vintage_0] = group_impacts[Vintages[1]] * np.nan

# Convert into a Pandas DataFrame, and fill in missing entries
# with zeros (missing entries happen when there were no updates
# for a given group in a given vintage)
group_impacts = (
    pd.concat(group_impacts, axis=1)
      .fillna(0)
      .reindex(group_counts.index).T)
group_impacts.index = forecasts.index

# Print the table of impacts from data in each group,
# along with a row with the "Total" impact
(group_impacts.T
    .append(group_impacts.sum(axis=1).rename('Total impact on 2020Q2 forecast'))
    .round(2).iloc[:, 1:])

if Dibujar_graficos==1:
    with sns.color_palette('deep'):
        fig, ax = plt.subplots(figsize=(14, 6))
    
        # Stacked bar plot showing the impacts by group
        group_impacts.plot(kind='bar', stacked=True, width=0.3, zorder=2, ax=ax);
    
        # Line plot showing the forecast for real GDP growth in 2020Q2 for each vintage
        x = np.arange(len(forecasts))
        ax.plot(x, forecasts, marker='o', color='k', markersize=7, linewidth=2)
        ax.hlines(0, -1, len(group_impacts) + 1, linewidth=1)
    
        # x-ticks
        labels = group_impacts.index.strftime('%b')
        ax.xaxis.set_ticklabels(labels)
        ax.xaxis.set_tick_params(size=0)
        ax.xaxis.set_tick_params(labelrotation='auto', labelsize=13)
    
        # y-ticks
        ax.yaxis.set_tick_params(direction='in', size=0, labelsize=13)
        ax.yaxis.grid(zorder=0)
        
        # title, remove spines
        ax.set_title('Evolution of real GDP growth nowcast: 2020Q2', fontsize=16, fontweight=600, loc='left')
        [ax.spines[spine].set_visible(False)
         for spine in ['top', 'left', 'bottom', 'right']]
        
        # base forecast vs updates
        ylim = ax.get_ylim()
        ax.vlines(0.5, ylim[0], ylim[1] + 5, linestyles='--')
        ax.annotate('Base forecast', (-0.2, 22), fontsize=14)
        ax.annotate(r'Updated forecasts and impacts from the "news" $\rightarrow$', (0.65, 22), fontsize=14)
    
        # legend
        ax.legend(loc='upper center', ncol=4, fontsize=13, bbox_to_anchor=(0.5, -0.1), frameon=False)
    
        fig.tight_layout();

"""
"""
# Construct the dynamic factor model
model = sm.tsa.DynamicFactorMQ(endog_m, endog_quarterly=endog_q,factors=1, factor_orders=2)
print(model.summary())
results = model.fit(disp=10)
print(results.summary())

# Get estimates of the global and labor market factors,
# conditional on the full dataset ("smoothed")
factor_names = ['Global']
mean = results.factors.smoothed
#mean=mean.rename(columns={1:'Global'})

with sns.color_palette('deep'):
    fig, ax = plt.subplots(figsize=(14, 3))
    mean.plot(ax=ax)
   
    ax.set(title='Estimated factors: smoothed estimates and 95% confidence intervals')
    fig.tight_layout();
    
    
rsquared = results.get_coefficients_of_determination(method='individual')


# Create point forecasts, 3 steps ahead
point_forecasts = results.forecast(steps=3)

# Print the forecasts for the first 5 observed variables
print(point_forecasts.T.head())


prediction_results = results.get_prediction(start='2000', end='2022')

point_predictions = prediction_results.predicted_mean


print(results.forecast(steps=5))
fore=res.forecast(steps=5)


model = sm.tsa.DynamicFactorMQ(
    endog_m, endog_quarterly=endog_q,
    factors=factors, factor_orders=factor_orders,
    factor_multiplicities=factor_multiplicities)







### Agrupar variables por factores
# Get the mapping of variable id to group name, for monthly variables
groups = defn_m[['description', 'group']].copy()

# Re-order the variables according to the definition CSV file
# (which is ordered by group)
columns = [name for name in defn_m['description']
           if name in dta['2020-02'].dta_m.columns]
for date in dta.keys():
    dta[date].dta_m = dta[date].dta_m.reindex(columns, axis=1)

# Add real GDP (our quarterly variable) into the "Output and Income" group
gdp_description = defn_q.loc['GDPC1', 'description']
groups = groups.append({'description': gdp_description, 'group': 'Output and Income'},
                       ignore_index=True)

# Display the number of variables in each group
(groups.groupby(factor, sort=False).count())
       .rename({'description': '# series in group'}, axis=1))



idx = orig_m.fecha
orig_m.index=idx


#orig_m.drop('sasdate', axis=1, inplace=True)
idx = orig_m.fecha
Datos.index=idx

# 4. Apply the transformations
for i in orig_m.columns:
    dta_m[i] = np.log(orig_m[i]).diff()

dta_m = orig_m.apply(transform, axis=0,
                       transforms=transform_m)










data=data.drop([0])
data.index = pd.period_range(start=data.loc[1,'fecha'], end=data.loc[data.index[-1],'fecha'], freq='M')
endog=data.drop(['fecha','SERIES'],axis=1)
endog_m = np.log(endog).diff().iloc[1:]

data_q = Book.sheets('QUARTERLY').range('A1')\
        .options(pd.DataFrame,expand='table',index=False,decimal='.').value
data_q=data_q.drop([0])
data_q.index = pd.period_range(start=data_q.loc[1,'fecha'], end=data_q.loc[data_q.index[-1],'fecha'], freq='q')
endog_q=data_q.drop(['fecha','SERIES'],axis=1)
endog_q = np.log(endog_q).diff().iloc[1:]




def load_fredmd_data(vintage):
    base_url = 'https://files.stlouisfed.org/files/htdocs/fred-md/'
    
    # - FRED-MD --------------------------------------------------------------
    # 1. Download data
    orig_m = (pd.read_csv(f'{base_url}/monthly/{vintage}.csv')
                .dropna(how='all'))
    
    # 2. Extract transformation information
    transform_m = orig_m.iloc[0, 1:]
    orig_m = orig_m.iloc[1:]

    # 3. Extract the date as an index
    orig_m.index = pd.PeriodIndex(orig_m.sasdate.tolist(), freq='M')
    orig_m.drop('sasdate', axis=1, inplace=True)

    # 4. Apply the transformations
    dta_m = orig_m.apply(transform, axis=0,
                         transforms=transform_m)

    # 5. Remove outliers (but not in 2020)
    dta_m.loc[:'2019-12'] = remove_outliers(dta_m.loc[:'2019-12'])

    # - FRED-QD --------------------------------------------------------------
    # 1. Download data
    orig_q = (pd.read_csv(f'{base_url}/quarterly/{vintage}.csv')
                .dropna(how='all'))

    # 2. Extract factors and transformation information
    factors_q = orig_q.iloc[0, 1:]
    transform_q = orig_q.iloc[1, 1:]
    orig_q = orig_q.iloc[2:]

    # 3. Extract the date as an index
    orig_q.index = pd.PeriodIndex(orig_q.sasdate.tolist(), freq='Q')
    orig_q.drop('sasdate', axis=1, inplace=True)

    # 4. Apply the transformations
    dta_q = orig_q.apply(transform, axis=0,
                          transforms=transform_q)

    # 5. Remove outliers (but not in 2020)
    dta_q.loc[:'2019Q4'] = remove_outliers(dta_q.loc[:'2019Q4'])
    
    # - Output datasets ------------------------------------------------------
    return types.SimpleNamespace(
        orig_m=orig_m, orig_q=orig_q,
        dta_m=dta_m, transform_m=transform_m,
        dta_q=dta_q, transform_q=transform_q, factors_q=factors_q)


# Load the vintages of data from FRED
dta = {date: load_fredmd_data(date)
       for date in ['2020-02', '2020-03', '2020-04', '2020-05', '2020-06']}


# Print some information about the base dataset
n, k = dta['2020-02'].dta_m.shape
start = dta['2020-02'].dta_m.index[0]
end = dta['2020-02'].dta_m.index[-1]

print(f'For vintage 2020-02, there are {k} series and {n} observations,'
      f' over the period {start} to {end}.')

with sns.color_palette('deep'):
    fig, axes = plt.subplots(3, figsize=(14, 6))

    # Plot the raw data from the February 2020 vintage, for:
    # 
    vintage = '2020-02'
    variable = 'RPI'
    start = '2000-01'
    end = '2020-01'

    # 1. Plot the original dataset, for 2000-01 through 2020-01
    dta[vintage].orig_m.loc[start:end, variable].plot(ax=axes[0])
    axes[0].set(title='Original data', xlim=('2000','2020'), ylabel='Billons of $')

    # 2. Plot the transformed data, still including outliers
    # (we only stored the transformation with outliers removed, so
    # here we'll manually perform the transformation)
    transformed = transform(dta[vintage].orig_m[variable],
                            dta[vintage].transform_m)
    transformed.loc[start:end].plot(ax=axes[1])
    mean = transformed.mean()
    iqr = transformed.quantile([0.25, 0.75]).diff().iloc[1]
    axes[1].hlines([mean - 10 * iqr, mean + 10 * iqr],
                   transformed.index[0], transformed.index[-1],
                   linestyles='--', linewidth=1)
    axes[1].set(title='Transformed data, with bands showing outliers cutoffs',
                xlim=('2000','2020'), ylim=(mean - 15 * iqr, mean + 15 * iqr),
                ylabel='Percent')
    axes[1].annotate('Outlier', xy=('2013-01', transformed.loc['2013-01']),
                     xytext=('2014-01', -5.3), textcoords='data',
                     arrowprops=dict(arrowstyle="->", connectionstyle="arc3"),)

    # 3. Plot the transformed data, with outliers removed (see missing value for 2013-01)
    dta[vintage].dta_m.loc[start:end, 'RPI'].plot(ax=axes[2])
    axes[2].set(title='Transformed data, with outliers removed',
                xlim=('2000','2020'), ylabel='Percent')
    axes[2].annotate('Missing value in place of outlier', xy=('2013-01', -1),
                     xytext=('2014-01', -2), textcoords='data',
                     arrowprops=dict(arrowstyle="->", connectionstyle="arc3"))
    
    fig.suptitle('Real Personal Income (RPI)',
                 fontsize=12, fontweight=600)

    fig.tight_layout(rect=[0, 0.00, 1, 0.95]);



# Definitions from the Appendix for FRED-MD variables
defn_m = pd.read_csv('data/fredmd_definitions.csv')
defn_m.index = defn_m.fred

# Definitions from the Appendix for FRED-QD variables
defn_q = pd.read_csv('data/fredqd_definitions.csv')
defn_q.index = defn_q.fred

# Example of the information in these files:
defn_m.head()












path = 'R:/_BankAnalytics/Centros/CoyEcono/Datos/publico/Modelizacion/Factor model/DATA/'
path += 'dATA.xlsX'
Book = xw.Book(path)
### Lectura de datos
Datos_M = Book.sheets('mensual').range('b2')\
    .options(pd.DataFrame,expand='table',index=False).value

Datos_Q = Book.sheets('trimestral').range('b2')\
    .options(pd.DataFrame,expand='table',index=False).value
    
Datos_M.index = pd.period_range(start=Datos_M.iloc[0,0], end=Datos_M.iloc[-1,0], freq='M')
Datos_Q.index = pd.period_range(start=Datos_Q.iloc[0,0], end=Datos_Q.iloc[-1,0], freq='Q')

 
### TRANSFORMACION DATA, estacionariedad ###

for i in ['afi','wage','eurostoxx']:
    Datos_M[i]=np.log(Datos_M[i]).diff(3)
    
#for i in ['ici','icc','icm']:
    #Datos_M[i]=np.log(Datos_M[i])
    
for i in ['PIB']:
    Datos_Q[i]=np.log(Datos_Q[i]).diff()
    
     
#Datos_M=Datos_M.loc['1995-04-01':'2019-01-01',:]
#Datos_Q=Datos_Q.loc['1995-04-01':'2019-01-01',:]

Datos_M=Datos_M.drop('fecha',axis=1)
Datos_Q=Datos_Q.drop('fecha',axis=1)

### MODEL ###

model = sm.tsa.DynamicFactorMQ(endog=Datos_M)
model = sm.tsa.DynamicFactorMQ(Datos_M, endog_quarterly=Datos_Q)
#model = sm.tsa.DynamicFactorMQ(Datos_M, endog_quarterly=Datos_Q, factors=factors, factor_orders=factor_orders, factor_multiplicities=factor_multiplicities)
model = sm.tsa.DynamicFactorMQ(Datos_M, endog_quarterly=Datos_Q, factors=2, factor_orders=2, factor_multiplicities=0)

print(model.summary())
results = model.fit(disp=10)
print(results.summary())




### POST ESTIMATION ###
# Estimacion de los factores (smoothed es utilizando toda la base de datos)
factor_names=['0','1']
mean = results.factors.smoothed[factor_names]

# Compute 95% confidence intervals
from scipy.stats import norm
std = pd.concat([results.factors.smoothed_cov.loc[name, name]
                 for name in factor_names], axis=1)
crit = norm.ppf(1 - 0.05 / 2)
lower = mean - crit * std
upper = mean + crit * std

with sns.color_palette('deep'):
    fig, ax = plt.subplots(figsize=(14, 3))
    mean.plot(ax=ax)
    
    for name in factor_names:
        ax.fill_between(mean.index, lower[name], upper[name], alpha=0.3)
    
    ax.set(title='Estimated factors: smoothed estimates and 95% confidence intervals')
    fig.tight_layout();


rsquared = results.get_coefficients_of_determination(method='individual')

top_ten = []
for factor_name in rsquared.columns[:2]:
    top_factor = (rsquared[factor_name].sort_values(ascending=False)
                                       .iloc[:10].round(2).reset_index())
    top_factor.columns = pd.MultiIndex.from_product([
        [f'Top ten variables explained by {factor_name}'],
        ['Variable', r'$R^2$']])
    top_ten.append(top_factor)
pd.concat(top_ten, axis=1)


print(res.forecast(steps=12))
print(res.impulse_responses(steps=5))
"""
# Detener el contador de tiempo y calcular el tiempo transcurrido
elapsed_time = time.time() - start_time
print(f"Tiempo de ejecución : {elapsed_time:.2f} segundos")
