CREATE TRIGGER UPMENUUSER ON dbo.XUSERPROFILE 
AFTER  UPDATE, INSERT
AS
  
DECLARE @userid AS VARCHAR
DECLARE @npecode AS VARCHAR
DECLARE @menuleve AS VARCHAR

BEGIN
	SELECT 
		@userid = usr_userid,
		@npecode = usr_npecode,
		@menuleve = usr_menuleve 
	FROM 
		INSERTED
END
  

IF UPDATE (usr_menuleve) 
BEGIN
	DELETE FROM MENUUSER 
	WHERE( 
		mu_npecode = @npecode 
		AND
		mu_userid = @userid
	)

	INSERT INTO MENUUSER  ( 
		mu_npecode, mu_userid,  mu_meopid, mu_accsflag, mu_accsflagread, mu_accsflagwrit
	)
        	SELECT 
		ma_npecode, @userid,  ma_meopid, ma_accsflag,  ma_accsflagread,ma_accsflagwrit
	FROM 
		MENUACCESS
	WHERE ( 
		ma_npecode = @npecode
		AND
		ma_melvid = @menuleve 
	)
	ORDER BY 
		ma_meopid            

END  

/*
IF NOT  ( UPDATE( usr_tbs ))  
BEGIN

        UPDATE XUSERPROFILE
        SET usr_tbs = 1
        WHERE usr_userid IN(SELECT usr_userid FROM INSERTED)
        AND usr_NPECODE IN(SELECT usr_NPECODE FROM INSERTED)

END 

        UPDATE XUSERPROFILE
        SET usr_modidate = GETDATE()
        WHERE usr_userid IN(SELECT usr_userid FROM INSERTED)
        AND usr_NPECODE IN(SELECT usr_NPECODE FROM INSERTED)
         

*/










