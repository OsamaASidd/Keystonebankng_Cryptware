-- SQL Script to add required columns to REPORT table if they don't exist

USE FIRSINVOICEDB;
GO

-- Check and add REPORTING_STATUS column
IF NOT EXISTS (SELECT * FROM sys.columns 
               WHERE object_id = OBJECT_ID('dbo.REPORT') 
               AND name = 'REPORTING_STATUS')
BEGIN
    ALTER TABLE dbo.REPORT
    ADD REPORTING_STATUS BIT DEFAULT 0;
    
    PRINT 'Added REPORTING_STATUS column';
END
ELSE
BEGIN
    PRINT 'REPORTING_STATUS column already exists';
END
GO

-- Check and add IRN column
IF NOT EXISTS (SELECT * FROM sys.columns 
               WHERE object_id = OBJECT_ID('dbo.REPORT') 
               AND name = 'IRN')
BEGIN
    ALTER TABLE dbo.REPORT
    ADD IRN NVARCHAR(500) NULL;
    
    PRINT 'Added IRN column';
END
ELSE
BEGIN
    PRINT 'IRN column already exists';
END
GO

-- Check and add QR_CODE column
IF NOT EXISTS (SELECT * FROM sys.columns 
               WHERE object_id = OBJECT_ID('dbo.REPORT') 
               AND name = 'QR_CODE')
BEGIN
    ALTER TABLE dbo.REPORT
    ADD QR_CODE NVARCHAR(MAX) NULL;
    
    PRINT 'Added QR_CODE column';
END
ELSE
BEGIN
    PRINT 'QR_CODE column already exists';
END
GO

-- Check and add LAST_UPDATED column
IF NOT EXISTS (SELECT * FROM sys.columns 
               WHERE object_id = OBJECT_ID('dbo.REPORT') 
               AND name = 'LAST_UPDATED')
BEGIN
    ALTER TABLE dbo.REPORT
    ADD LAST_UPDATED DATETIME NULL;
    
    PRINT 'Added LAST_UPDATED column';
END
ELSE
BEGIN
    PRINT 'LAST_UPDATED column already exists';
END
GO

-- Create index on REPORTING_STATUS for better query performance
IF NOT EXISTS (SELECT * FROM sys.indexes 
               WHERE name = 'IDX_REPORT_REPORTING_STATUS' 
               AND object_id = OBJECT_ID('dbo.REPORT'))
BEGIN
    CREATE INDEX IDX_REPORT_REPORTING_STATUS 
    ON dbo.REPORT(REPORTING_STATUS)
    WHERE REPORTING_STATUS = 0 OR REPORTING_STATUS IS NULL;
    
    PRINT 'Created index on REPORTING_STATUS';
END
ELSE
BEGIN
    PRINT 'Index on REPORTING_STATUS already exists';
END
GO

PRINT 'Database setup completed successfully';
