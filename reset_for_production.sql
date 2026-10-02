-- Run this script ONCE after updating config.json with the production API key.
-- It resets all PREPROD-processed records so the scheduler picks them up again
-- and resubmits them to the production API.
--
-- PREPROD_IRN and PREPROD_QR_CODE are preserved as a permanent audit trail.

UPDATE [FIRSINVOICEDB].[dbo].[REPORT]
SET REPORTING_STATUS = NULL,
    IRN              = NULL,
    QR_CODE          = NULL,
    VALIDATION_ERROR = NULL,
    LAST_UPDATED     = NULL,
    ENVIRONMENT      = NULL
WHERE ENVIRONMENT = 'PREPROD';

-- Confirm how many records were reset
SELECT 'Reset for production: ' + CAST(@@ROWCOUNT AS VARCHAR) + ' records';

-- Verify backup columns are intact
SELECT COUNT(*) AS PreprodRecordsWithBackup
FROM [FIRSINVOICEDB].[dbo].[REPORT]
WHERE PREPROD_IRN IS NOT NULL;
