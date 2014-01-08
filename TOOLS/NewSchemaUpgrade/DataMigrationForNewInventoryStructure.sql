/*-----------------------------------------------------------------------------


----------To check if logical warehouse are missing from invtissue/logware tables
select 
distinct iid_tologiware
from invtissuedetl
inner join invtissue on ii_trannumb = iid_trannumb and ii_npecode = iid_npecode 
inner join stockmaster on  stk_stcknumb= iid_stcknumb and stk_npecode = iid_npecode 
where iid_stcknumb in (select stk_stcknumb from stockmaster)
 and (iid_fromlogiware not in (select lw_code from logwar where lw_npecode = iid_npecode ) or
 iid_tologiware not in (select lw_code from logwar where lw_npecode = iid_npecode ))

----------To check if logical warehouse are missing from invtreceipt/logware tables
select 
ird_trannumb,	 ird_transerl,	 ird_npecode, invtreceipt.ir_trantype,	 ird_compcode,	
ird_fromlogiware,	 ird_tologiware, ird_ware, invtreceipt.ir_ware,	ird_fromsubloca,	 
ird_tosubloca, ird_ponumb,	 ird_liitnumb,	 ird_stcknumb, ird_newstcknumb,	 ird_origcond,	 ird_newcond,	
ird_stckdesc,	 ird_newdesc,	 ird_ps,	 null,	 ird_primqty,	 ird_secoqty,
	stk_primuon ,stk_secouom , ird_unitpric, ird_reprcost, ird_stcktype,	 ird_owle,	 ird_leasecomp, 
ird_creadate, ird_creauser, ird_modidate, ird_modiuser, ird_tbs
from invtreceiptdetl
inner join invtreceipt on ir_trannumb = ird_trannumb and ir_npecode = ird_npecode and ir_trantype <> 'R'
inner join stockmaster on  stk_stcknumb = ird_stcknumb and stk_npecode = ird_npecode
where ird_stcknumb in (select stk_stcknumb from stockmaster)
 and (ird_fromlogiware not in (select lw_code from logwar where lw_npecode = ird_npecode ) or
 ird_tologiware not in (select lw_code from logwar where lw_npecode = ird_npecode ))



-----------INSERTING SOME MISSING LOCATIONS/logwars which are used by some of the transactions.
INSERT INTO LOCATION
SELECT * FROM OLDPECTEN..LOCATION WHERE LOC_LOCACODE IN ('LIP','GENERAL','D96')
UPDATE LOCATION SET LOC_ACTVFLAG = 0 WHERE  LOC_LOCACODE IN ('LIP','GENERAL','D96')

INSERT INTO LOCATION SELECT 'GENERAL', loc_npecode, loc_compcode, 'FOR IA-1', loc_adr1, loc_adr2, loc_city,loc_stat, loc_zipc, loc_ctry, '', '', '', '', loc_cont, loc_gender, 0, loc_tbs, loc_creadate,loc_creauser, loc_modidate ,loc_modiuser FROM LOCATION WHERE LOC_LOCACODE IN ('LIP')
insert into logwar select 'unique' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select '' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'm&T' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'DRL' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'CHM' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'PRD' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'SUR' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'
insert into logwar select 'C-20-C' ,   lw_npecode, ' Schema upgrade' ,                                 lw_tbs, lw_creadate,                                            lw_creauser,          lw_modidate ,                                           lw_modiuser,          0  from logwar where lw_code ='general'





*/



-----------Import from invtreceipt table ---------------------
insert into inventorytransac  
select ir_trannumb,ir_trantype, ir_npecode, ir_compcode, irr_remk, ird_curr, ird_currvalu, ir_creadate, ir_creauser, 
ir_modidate, ir_modiuser, ir_tbs from invtreceipt
LEFT  join invtreceiptrem on irr_trannumb = ir_trannumb and irr_npecode = ir_npecode and irr_linenumb =1
inner join invtreceiptdetl on  ird_trannumb = ir_trannumb and irD_npecode = iR_npecode and ird_transerl =1 


----------Only receipt trasactions ---------------------------
insert into inventory
(Transaction# ,	TransactionLine ,	Namespace ,	TransactionType ,	Company ,	
LogicalWarehouseFrom ,	LogicalWarehouseTo , LocationFrom ,	LocationTo ,	SubLocationFrom ,	
SubLocationTo , PO ,	POITEM,	StockNumberFrom ,	StockNumberTo ,	ConditionFrom ,	ConditionTo ,	
StockDescriptionFrom ,StockDescriptionTo ,PS ,Serial ,PrimaryQuantity ,SecondaryQuantity,
PrimaryUnit ,	SecondaryUnit ,	UnitPrice ,AdditionalCost,	StockType ,	OWLE ,	LeaseCompany ,CreaDate, CreaUser
,ModiDate, ModiUser, tbs)

select 
ird_trannumb,	 ird_transerl,	 ird_npecode, invtreceipt.ir_trantype,	 ird_compcode,	
ird_fromlogiware,	 ird_tologiware, ird_ware, invtreceipt.ir_ware,	ird_fromsubloca,	 
ird_tosubloca, ird_ponumb,	 ird_liitnumb,	 ird_stcknumb, ird_newstcknumb,	 ird_origcond,	 ird_newcond,	
ird_stckdesc,	 ird_newdesc,	 ird_ps,	 null,	 ird_primqty,	 ird_secoqty,
	stk_primuon ,stk_secouom , ird_unitpric, ird_reprcost, ird_stcktype,	 ird_owle,	 ird_leasecomp, 
ird_creadate, ird_creauser, ird_modidate, ird_modiuser, ird_tbs
from invtreceiptdetl
inner join invtreceipt on ir_trannumb = ird_trannumb and ir_npecode = ird_npecode and ir_trantype <> 'R'
inner join stockmaster on  stk_stcknumb = ird_stcknumb and stk_npecode = ird_npecode

----------All other trasactions ---------------------------
insert into inventory
(Transaction# ,	TransactionLine ,	Namespace ,	TransactionType ,	Company ,	
LogicalWarehouseFrom ,	LogicalWarehouseTo , LocationFrom ,	LocationTo ,	SubLocationFrom ,	
SubLocationTo , PO ,	POITEM,	StockNumberFrom ,	StockNumberTo ,	ConditionFrom ,	ConditionTo ,	
StockDescriptionFrom ,StockDescriptionTo ,PS ,Serial ,PrimaryQuantity ,SecondaryQuantity,
PrimaryUnit ,	SecondaryUnit ,	UnitPrice ,AdditionalCost,	StockType ,	OWLE ,	LeaseCompany ,CreaDate, CreaUser
,ModiDate, ModiUser, tbs)

select 
ird_trannumb,	 ird_transerl,	 ird_npecode, invtreceipt.ir_trantype,	 ird_compcode,	
ird_fromlogiware,	 ird_tologiware, ird_ware, invtreceipt.ir_ware,	ird_fromsubloca,	 
ird_tosubloca, ird_ponumb,	 ird_liitnumb,	 ird_stcknumb, ird_newstcknumb,	 ird_origcond,	 ird_newcond,	
ird_stckdesc,	 ird_newdesc,	 ird_ps,	 null,	 ird_primqty,	 ird_secoqty,
	poi_primuom ,poi_secouom , ird_unitpric, ird_reprcost, ird_stcktype,	 ird_owle,	 ird_leasecomp, 
ird_creadate, ird_creauser, ird_modidate, ird_modiuser, ird_tbs
from invtreceiptdetl
inner join invtreceipt on ir_trannumb = ird_trannumb and ir_npecode = ird_npecode and ir_trantype = 'R'
inner join poitem on  poi_ponumb = ird_ponumb and poi_comm= ird_stcknumb and poi_npecode = ird_npecode and poi_liitnumb =ird_liitnumb


-------- for Invtissue table ------
insert into inventorytransac  
select ii_trannumb,ii_trantype, ii_npecode, ii_compcode, iir_remk, iid_curr, iid_currvalu, ii_creadate, ii_creauser, 
ii_modidate, ii_modiuser, ii_tbs from invtissue
LEFT  join invtissuerem on iir_trannumb = ii_trannumb and iir_npecode = ii_npecode and iir_linenumb =1
inner join invtissuedetl on  iid_trannumb = ii_trannumb and iiD_npecode = ii_npecode and iid_transerl =1 

-------- for Invtissue table/ stocks not in stockmaster ------
insert into inventory
(Transaction# ,	TransactionLine ,	Namespace ,	TransactionType ,	Company ,	
LogicalWarehouseFrom ,	LogicalWarehouseTo , LocationFrom ,	LocationTo ,	SubLocationFrom ,	
SubLocationTo , PO ,	POITEM,	StockNumberFrom ,	StockNumberTo ,	ConditionFrom ,	ConditionTo ,	
StockDescriptionFrom ,StockDescriptionTo ,PS ,Serial ,PrimaryQuantity ,SecondaryQuantity,
PrimaryUnit ,	SecondaryUnit ,	UnitPrice ,AdditionalCost,	StockType ,	OWLE ,	LeaseCompany ,CreaDate, CreaUser
,ModiDate, ModiUser, tbs)
select 
iid_trannumb,	 iid_transerl,	 iid_npecode, invtissue.ii_trantype,	 iid_compcode,	
iid_fromlogiware,	 iid_tologiware, iid_ware, invtissue.ii_ware,	iid_fromsubloca,	 
iid_tosubloca, iid_ponumb,	 iid_liitnumb,	 iid_stcknumb, null,	 iid_origcond,	 iid_newcond,	
iid_stckdesc,	 null,	 iid_ps,	 null,	 iid_primqty,	 iid_secoqty,
	poi_primuom ,poi_secouom , iid_unitpric, null, iid_stcktype,	 iid_owle,	 iid_leasecomp, 
iid_creadate, iid_creauser, iid_modidate, iid_modiuser, iid_tbs
from invtissuedetl
inner join invtissue on ii_trannumb = iid_trannumb and ii_npecode = iid_npecode 
inner join poitem on  poi_comm= iid_stcknumb and poi_npecode = iid_npecode and poi_liitnumb =1
where iid_stcknumb not in (select stk_stcknumb from stockmaster)



-------- for Invtissue table/ stocks in stockmaster ------
insert into inventory
(Transaction# ,	TransactionLine ,	Namespace ,	TransactionType ,	Company ,	
LogicalWarehouseFrom ,	LogicalWarehouseTo , LocationFrom ,	LocationTo ,	SubLocationFrom ,	
SubLocationTo , PO ,	POITEM,	StockNumberFrom ,	StockNumberTo ,	ConditionFrom ,	ConditionTo ,	
StockDescriptionFrom ,StockDescriptionTo ,PS ,Serial ,PrimaryQuantity ,SecondaryQuantity,
PrimaryUnit ,	SecondaryUnit ,	UnitPrice ,AdditionalCost,	StockType ,	OWLE ,	LeaseCompany ,CreaDate, CreaUser
,ModiDate, ModiUser, tbs)
select 
iid_trannumb,	 iid_transerl,	 iid_npecode, invtissue.ii_trantype,	 iid_compcode,	
iid_fromlogiware,	 iid_tologiware, iid_ware, invtissue.ii_ware,	iid_fromsubloca,	 
iid_tosubloca, iid_ponumb,	 iid_liitnumb,	 iid_stcknumb, null,	 iid_origcond,	 iid_newcond,	
iid_stckdesc,	 null,	 iid_ps,	 null,	 iid_primqty,	 iid_secoqty,
	stk_primuon ,stk_secouom , iid_unitpric, null, iid_stcktype,	 iid_owle,	 iid_leasecomp, 
iid_creadate, iid_creauser, iid_modidate, iid_modiuser, iid_tbs
from invtissuedetl
inner join invtissue on ii_trannumb = iid_trannumb and ii_npecode = iid_npecode 
inner join stockmaster on  stk_stcknumb= iid_stcknumb and stk_npecode = iid_npecode 
where iid_stcknumb in (select stk_stcknumb from stockmaster)

