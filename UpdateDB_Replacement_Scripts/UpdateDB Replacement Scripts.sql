
-- run this only if DB is configured for UPdate DB
If (@@UpdateDBActive = 1)

BEGIN

		--Get all those updates from Inserted table which belong to 'Houst' namespace
		Declare cursor cur_supplier local
		for
		select sup_code,   sup_npecode, sup_flag, sup_name, sup_adr1, sup_adr2, sup_city, sup_stat, sup_zipc, sup_ctry, sup_phonnumb, sup_faxnumb, sup_telxnumb, sup_contaname,
		sup_contaph, sup_contaFax, sup_mail, sup_actvflag, sup_sapcode, sup_remk, sup_tbs, sup_creadate, sup_creauser, sup_modidate, sup_modiuser
		from inserted where sup_npecode ='Houst'

		OPEN cur_supplier

		FETCH cur_supplier INTO
		@sup_code,   @sup_npecode, @sup_flag, @sup_name, @sup_adr1, @sup_adr2, @sup_city, @sup_stat, @sup_zipc, @sup_ctry, @sup_phonnumb, @sup_faxnumb, @sup_telxnumb, @sup_contaname,
		@sup_contaph, @sup_contaFax, @sup_mail, @sup_actvflag, @sup_sapcode, @sup_remk, @sup_tbs, @sup_creadate, @sup_creauser, @sup_modidate, @sup_modiuser

		WHILE @@fetchstatus >0
		BEGIN

					-- insert the supplier in the namespaces in which it does not exist
					insert into supplier
					(sup_code,   sup_npecode, sup_flag, sup_name, sup_adr1, sup_adr2, sup_city, sup_stat, sup_zipc, sup_ctry, sup_phonnumb, sup_faxnumb, 
					sup_telxnumb, sup_contaname,
					sup_contaph, sup_contaFax, sup_mail, sup_actvflag, sup_sapcode, sup_remk, sup_tbs, sup_creadate, sup_creauser, sup_modidate, sup_modiuser)
					VALUES
					(select @sup_code, npce_code, @sup_flag, @sup_name, @sup_adr1, @sup_adr2, @sup_city, @sup_stat, @sup_zipc, @sup_ctry, @sup_phonnumb, @sup_faxnumb, 
					@sup_telxnumb, @sup_contaname,
					@sup_contaph, @sup_contaFax, @sup_mail, @sup_actvflag, @sup_sapcode, @sup_remk, @sup_tbs, @sup_creadate, @sup_creauser, getdate(), @sup_modiuser)
					from namespace
					
					where npce_code not in 
					--get all the namespaces in which this supplier exists
					( select sup_npecode from supplier where sup_code = @sup_code)
					and
					-- making sure namepsace i which updates are being updated\ inserted is configured to receive updates
					npce_code in (select distinct dis_npecode from distribution)

					and
					sup_code=@sup_code

					--apply an update on all those namespaces in which it already exists
					update supplier
					set 
					--sup_code=@sup_code,   sup_npecode=@sup_npecode,
					sup_flag=@sup_flag, sup_name=@sup_name, sup_adr1=@sup_adr1, sup_adr2=@sup_adr2, sup_city=@sup_city, 
					sup_stat=@sup_stat, sup_zipc=@sup_zipc, sup_ctry=@sup_ctry, sup_phonnumb=@sup_phonnumb, sup_faxnumb=@sup_faxnumb, sup_telxnumb=@sup_telxnumb, 
					sup_contaname=@sup_contaname,
					sup_contaph=@sup_contaph, sup_contaFax=@sup_contaFax, sup_mail=@sup_mail, sup_actvflag=@sup_actvflag, sup_sapcode=@sup_sapcode, 
					sup_remk=@sup_remk, sup_tbs=@sup_tbs, sup_creadate=@sup_creadate, sup_creauser=@sup_creauser, sup_modidate=getdate(), sup_modiuser=@sup_modiuser
					where npce_code in 
					--get all the namespaces in which this supplier exists
					( select sup_npecode from supplier where sup_code = @sup_code and sup_npecode <>'houst')
					and
					-- making sure namepsace i which updates are being updated\ inserted is configured to receive updates
					npce_code in (select distinct dis_npecode from distribution)
					and 
					sup_code=@sup_code

					FETCH cur_supplier INTO
					@sup_code,   @sup_npecode, @sup_flag, @sup_name, @sup_adr1, @sup_adr2, @sup_city, @sup_stat, @sup_zipc, @sup_ctry, @sup_phonnumb, @sup_faxnumb, @sup_telxnumb, @sup_contaname,
					@sup_contaph, @sup_contaFax, @sup_mail, @sup_actvflag, @sup_sapcode, @sup_remk, @sup_tbs, @sup_creadate, @sup_creauser, @sup_modidate, @sup_modiuser


					--get a list of all the namespaces in which this supplier code does not exist



					/*(select sup_code,   sup_npecode, sup_flag, sup_name, sup_adr1, sup_adr2, sup_city, sup_stat, sup_zipc, sup_ctry, sup_phonnumb, sup_faxnumb, sup_telxnumb, sup_contaname,
					sup_contaph, sup_contaFax, sup_mail, sup_actvflag, sup_sapcode, sup_remk, sup_tbs, sup_creadate, sup_creauser, sup_modidate, sup_modiuser from inserted)

					where 
					*/
		END


END