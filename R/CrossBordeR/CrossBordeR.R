# Objective: to write a series of functions that allow easy manipulation of
# common troubleshooting operations/reporting done in SI databases

# List of common operations:

# 1. Joining fund- & port-level information
# 2. Comparing OriginationMaster, CleanMaster, CleanGlobalMaster, & Measures
# 3. Joining CB and GL information (in combination with #1)
# TODO: roll CM & M measures to Fund level (CGM >> CM & M from double-counting)


require(RODBC)
require(tidyr)
require(dplyr)
require(lubridate)
require(ggvis)


get.data <- function(db = 'GLDB', Id, start.period, end.period = NULL) {
  con <- odbcConnect(toupper(db))

  # Select columns
  columns <- paste('part.ParticipantName, pd.PeriodId, p.PortfolioId,
                   cgm.FundId,', 'mkt.OriginationMarket,',
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
         'LEFT JOIN CrossBorder.dbo.CleanMaster cm ON cgm.FundId = cm.FundId AND cgm.PeriodId = cm.PeriodId',
         'LEFT JOIN siGlobalResearch.dbo.vFund f ON cgm.FundId = f.FundId',
         'LEFT JOIN siGlobalResearch.dbo.vPortfolio p ON f.PortfolioId = p.PortfolioId',
         'LEFT JOIN Vendors.CBSO.OriginationMarkets mkt ON cm.OrigMarketId = mkt.MarketId',
         'LEFT JOIN Vendors.CBSO.Participant part ON cm.ParticipantId = part.ParticipantId',
         'LEFT JOIN siGlobalResearch.dbo.siPeriod pd ON cgm.PeriodId = pd.PeriodId',
       build.conditional(Id, start.period, end.period)
    )
  return(query)
  # Retrieve data from DB
  cgm.query <- tbl_df(sqlQuery(channel = con, query = query))

  # Add CrossBorder.dbo.Measures
  m.query <- find.measures(Id, start.period, con, end.period)

  # Join m.query to out
  out <- m.query %>% left_join(cgm.query)

  # Ensure PeriodId is POSIX
  out$PeriodId <- parse_date_time(out$PeriodId, 'Y!m*!')

  # Close connections
  odbcCloseAll()

  return(out)
}


find.measures <- function(Id, start.period, db, end.period) {
  columns <- c('ParticipantName', 'PeriodId', 'PortfolioId', 'FundId', 'OriginationMarket',
               'Local_FundMarket', 'DataSourceId', 'SalesType', 'SalesChannel',
               '[1] / 1000000 AS M_Assets', '[2] / 1000000 AS M_Gross',
               '[3] / 1000000 AS M_Net', '[4] / 1000000 AS M_Redemp')
  inner.cols <- c('part.ParticipantName', 'pd.PeriodId', 'b.PortfolioId', 'a.FundId',
                  'mkt.OriginationMarket', 'a.Local_FundMarket', 'a.DataSourceId',
                  'st.SalesType', 'sc.SalesChannel', 'a.MeasureTypeId', 'a.Value')

  # Join the column vectors
  columns <- paste(columns, collapse = ', ')
  inner.cols <- paste(inner.cols, collapse = ', ')

  # build query
  query <- paste(
    'SELECT', columns,
      'FROM (',
        'SELECT', inner.cols,
          'FROM CrossBorder.dbo.Measures a',
            'LEFT JOIN siGlobalResearch.dbo.vFund b ON b.FundId = a.FundId',
            'LEFT JOIN Vendors.CBSO.OriginationMarkets mkt ON a.OrigMarketId = mkt.MarketId',
            'LEFT JOIN CrossBorder.dbo.DistributionType sc ON a.SalesChannelId = sc.SalesId',
            'LEFT JOIN CrossBorder.dbo.DistributionType st ON a.SalesTypeId = st.SalesId',
            'LEFT JOIN siGlobalResearch.dbo.vFund f ON a.FundId = f.FundId',
            'LEFT JOIN siGlobalResearch.dbo.vPortfolio p ON f.PortfolioId = p.PortfolioId',
            'LEFT JOIN Vendors.CBSO.Participant part ON p.ManagerId = part.ManagerId',
            'LEFT JOIN siGlobalResearch.dbo.siPeriod pd ON a.PeriodId = pd.PeriodId',
          build.conditional(Id, start.period, end.period),
      ') m',
    'PIVOT',
      '(SUM(m.Value) FOR m.MeasureTypeId IN ([1], [2], [3], [4])) piv'
    )

  # how to deal with cgm.PeriodId in build.conditional() when called
  # from find.measures (need m.PeriodId)
  tbl_df(sqlQuery(channel = db, query = query))
}


build.conditional <- function(Id, start.period, end.period) {
  # manager vs portfolio
  if (class(Id) == 'character') {
    out <- paste("WHERE part.ParticipantName ='", Id, "'", sep = '')
  } else {
    out <- paste('WHERE p.PortfolioId =', Id)
  }

  # single period vs interval
  if (class(end.period) == 'NULL') {
    out <- paste(out, 'AND pd.PeriodId =', start.period)
  } else {
    out <- paste(out, 'AND pd.PeriodId BETWEEN', start.period, 'AND',
                 end.period)
  }

  return(out)
}
