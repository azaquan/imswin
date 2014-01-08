CREATE TRIGGER [UPDATE_MENUUSER] ON dbo.MENUUSER
FOR INSERT, UPDATE AS

DECLARE @meopid AS VARCHAR(10)
DECLARE @oaccsflag AS BIT
DECLARE @accsflag AS BIT
DECLARE @oaccsflagread AS BIT
DECLARE @accsflagread AS BIT
DECLARE @oaccsflagwrit AS BIT
DECLARE @accsflagwrit AS BIT
DECLARE @npecode AS [NPECODE]
DECLARE @userid AS VARCHAR(15)
DECLARE @modiuserid AS VARCHAR(15)
DECLARE @lastDate AS DATETIME --added by Juan Gonzalez 2007/8/4
DECLARE @modiuser AS VARCHAR(15) --added by Juan Gonzalez 2007/8/4

BEGIN
	UPDATE 
		MENUUSER
	SET 
		mu_MODIDATE = GETDATE(), 
		mu_MODIUSER = USER_NAME()
	WHERE 
		mu_meopid IN(
			SELECT mu_meopid FROM INSERTED)
	AND 
		mu_NPECODE IN(
			SELECT mu_npecode FROM INSERTED)
	AND mu_userid  IN(
		SELECT mu_userid FROM INSERTED)
	
	IF NOT  ( UPDATE( mu_tbs ))  BEGIN
		UPDATE 
			MENUUSER
		SET 
			mu_tbs = 1
		WHERE 
			mu_meopid IN(
				SELECT mu_meopid FROM INSERTED)
		AND 
			mu_NPECODE IN(
				SELECT mu_NPECODE FROM INSERTED) 
		AND 
			mu_userid in (
				Select mu_userid from INSERTED)
	END 

--Added by Juan Gonzalez 2007/06/21
BEGIN
	SELECT
		@meopid=mu_meopid,
		@oaccsflag=mu_accsflag,
		@oaccsflagread=mu_accsflagread,
		@oaccsflagwrit=mu_accsflagwrit			
	FROM
		DELETED
		
	SELECT
		@userid=mu_userid,
		@npecode=mu_npecode,
		@meopid=mu_meopid,
		@accsflag=mu_accsflag,
		@accsflagread=mu_accsflagread,
		@accsflagwrit=mu_accsflagwrit,	
		@lastDate=mu_modidate,
		@modiuserid=mu_modiuser
	FROM
		INSERTED				
	
	SELECT
		@modiuser=eve_from
	FROM
		XEVENT
	WHERE
		DATEDIFF(SECOND,eve_modidate,@lastDate)<=1
		
	IF @modiuser IS NOT NULL 
		SELECT @modiuserid=@modiuser
				
    INSERT INTO
    		MENUUSERHISTORY
	(
		muh_userid,
		muh_npecode,
		muh_meopid,
		muh_oaccsflag,
		muh_accsflag,
		muh_oaccsflagread,
		muh_accsflagread,		
		muh_oaccsflagwrit,
		muh_accsflagwrit,
		muh_tbs,
		muh_modiuser)		
	VALUES(
		@userid,
		@npecode,
		@meopid,
		@oaccsflag,
		@accsflag,
		@oaccsflagread,
		@accsflagread,		
		@oaccsflagwrit,
		@accsflagwrit,
		1,
		@modiuserid
		)
	END
END







