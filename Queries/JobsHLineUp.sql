-- Query Name: JobsHLineUp
-- Type: Select
-- Date Exported: 2/21/2026 9:53:55 AM
-- ----------------------------------------------------------------------

SELECT AllLiveJobs.StockCode, [1stJobs].FirstOfJob, [1stJobs].FirstOfQtyToMake, [1stJobs].FirstOfJobDeliveryDate, [2ndJobs].FirstOfJob1, [2ndJobs].FirstOfQtyToMake, [2ndJobs].FirstOfJobDeliveryDate, [3rdJobs].FirstOfJob, [3rdJobs].FirstOfQtyToMake, [3rdJobs].FirstOfJobDeliveryDate
FROM ((AllLiveJobs LEFT JOIN 1stJobs ON AllLiveJobs.StockCode=[1stJobs].StockCode) LEFT JOIN 2ndJobs ON AllLiveJobs.StockCode=[2ndJobs].StockCode) LEFT JOIN 3rdJobs ON AllLiveJobs.StockCode=[3rdJobs].StockCode;

