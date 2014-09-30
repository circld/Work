# Objective: to write a series of functions that allow easy manipulation of
# common troubleshooting operations/reporting done in SI databases

# List of common operations:

# 1. Joining fund- & port-level information
# 2. Comparing OriginationMaster, CleanMaster, CleanGlobalMaster, & Measures
# 3. Joining CB and GL information (in combination with #1)

require('RODBC')
require('tidyr')
require('dplyr')


find.port <- function(port.num, period.id, db = 'GLDB') {
  con <- odbcConnect(toupper(db))

  # Select columns
  columns <- paste('p.PortfolioId, cgm.FundId, mkt.OriginationMarket,',
                   'cm.SalesChannel, cm.SalesType,',
                   'cm.Assets / 1000000 [CM_Assets],',
                   'cgm.Asset / 1000000 [CGM_Assets],',
                   'cm.NetSales / 1000000 [CM_Net],',
                   'cgm.NetSales / 1000000 [CGM_Net],',
                   'cm.GrossSales / 1000000 [CM_Gross]')
  # Build query
  query <- paste(
    'SELECT', columns,
       'FROM CrossBorder.dbo.CleanGlobalMaster cgm',
         'JOIN CrossBorder.dbo.CleanMaster cm ON cgm.FundId = cm.FundId AND cgm.PeriodID = cm.PeriodID',
         'JOIN siGlobalResearch.dbo.vFund f ON cgm.FundId = f.FundId',
         'JOIN siGlobalResearch.dbo.vPortfolio p ON f.PortfolioId = p.PortfolioId',
         'JOIN Vendors.CBSO.OriginationMarkets mkt ON cm.OrigMarketID = mkt.MarketID',
       'WHERE cgm.PeriodId =', period.id, 'AND p.PortfolioId =', port.num)

  # Retrieve data from DB
  cgm.query <- tbl_df(sqlQuery(channel = con, query = query))

  # Add CrossBorder.dbo.Measures
  m.query <- find.measures(port.num, period.id, con)

  # Join m.query to out
  out <- m.query %>% left_join(cgm.query)

  # Close connections
  odbcCloseAll()

  return(out)
}

find.measures <- function(port.num, period.id, db) {
  columns <- c('PortfolioID', 'FundId', 'OriginationMarket', 'Local_FundMarket',
               'DataSourceID', 'SalesType', 'SalesChannel',
               '[1] / 1000000 AS M_Assets', '[2] / 1000000 AS M_Gross',
               '[3] / 1000000 AS M_Net', '[4] / 1000000 AS M_Redemp')
  inner.cols <- c('b.PortfolioId', 'a.FundId', 'mkt.OriginationMarket',
                  'a.Local_FundMarket', 'a.DataSourceId', 'st.SalesType',
                  'sc.SalesChannel', 'a.MeasureTypeId', 'a.Value')

  # Join the column vectors
  columns <- paste(columns, collapse = ', ')
  inner.cols <- paste(inner.cols, collapse = ', ')

  # build query
  query <- paste(
    'SELECT', columns,
      'FROM (',
        'SELECT', inner.cols,
          'FROM CrossBorder.dbo.Measures a',
            'JOIN siGlobalResearch.dbo.vFund b ON b.FundID = a.FundID',
            'JOIN Vendors.CBSO.OriginationMarkets mkt ON a.OrigMarketID = mkt.MarketID',
            'JOIN CrossBorder.dbo.DistributionType sc ON a.SalesChannelID = sc.SalesID',
            'JOIN CrossBorder.dbo.DistributionType st ON a.SalesTypeID = st.SalesID',
          'WHERE b.PortfolioID =', port.num, 'AND a.PeriodId =', period.id,
      ') m',
    'PIVOT',
      '(SUM(m.Value) FOR m.MeasureTypeID IN ([1], [2], [3], [4])) piv'
    )

  tbl_df(sqlQuery(channel = db, query = query))
}
