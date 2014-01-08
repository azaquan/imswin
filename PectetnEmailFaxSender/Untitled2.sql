select * from po where po_ponumb like '%muz%'
select * from porec where porc_ponumb like '%muz%'
--delete from porec where porc_ponumb like 'muz-080808'

select * from po 
inner join poitem p on po.po_ponumb = p.poi_ponumb
where po_ponumb like '%muz%'
--delete  FROM EMAILFAX where rowid in( 87,88)
SELECT send, * FROM EMAILFAX  where send is null
SELECT * FROM EMAILFAXCONFIG
select * from emailfaxconfigerrors

update EMAILFAX set send=null

update EMAILFAX set attachmentfile = replace(attachmentfile,';','')
update po set po_stas = 'OH' where po_ponumb like 'A013204'
update po set po_stas = 'OH' where po_ponumb like '%muz%'

select * from pecten.dbo.porec where porc_ponumb like 'muz-test-020609'
/*
insert into porec
select x.po_ponumb porc_ponumb, y.porc_npecode , y.porc_recpnumb , y.porc_rec , y.porc_tbs ,porc_creadate  , y.porc_creauser , y.porc_modidate , y.porc_modiuser
 from po x, pecten.dbo.porec y where x.po_ponumb like '%muz%' and y.porc_ponumb like 'muz-test-020609'
order by po_ponumb*/