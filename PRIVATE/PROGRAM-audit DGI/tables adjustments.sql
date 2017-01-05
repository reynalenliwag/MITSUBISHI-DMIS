--alter table ALL_CUSTOMER_TERMS
--add terms_expiry_year int
--
--alter table ALL_CUSTOMER_TERMS
--add terms_expiry_month int 
--
--alter table ALL_CUSTOMER_TERMS
--add Credit_start smalldatetime 

update all_profile set version =0
--
--
SELECT * FROM ALL_CUSTOMER_TERMS where customercode = 'A00031'

select * from all_Customer WHERE cuscde = 'A00031'


select sum(amount) from cmis_off_dt where trantype = 'VI'

SELECT * FROM ALL_CUSTOMER_TERMS WHERE customercode in (select cuscde from all_customer where acctname = 'ALLIED BANK')


select * from SMIS_MRRINV_TABLE where customercode = '100001'
select * from SMIS_SALESORDER where vi_no = '000002'

select * from SMIS_MRRINV_TABLE  WHERE IGNKEY='31231231'
select * from SMIS_MRRINV_TABLE 
where   ISTATUS='A' and  PROSPECTID=0, PROSPECTCOUNTER=ISNULL(PROSPECTCOUNTER,0) + 1 , WITHPROSBUYERS='Y' , CUSTOMERCODE='100001' WHERE IGNKEY='31231231'


------------saving invoice

UPDATE SMIS_MRRINV_TABLE SET PROSPECTID=NULL, customercode=NULL,datereleased=null, invoiceddate=null,IStatus='O', Released=0, WithProsBuyers='N'  WHERE IGNKEY='31231231'

UPDATE   SMIS_MRRINV_TABLE SET Istatus='S',  CustomerCode = '100001' , ProspectID = 0 , Released = 0 ,VI_No='000001' WHERE ignkey='31231231'

------------query

select 
sum(baltofinanced)-(select isnull(sum(amount),0) 
from cmis_off_dt where trantype = 'VI' and paidna = 1 and cuscde = 'A00031')
 from SMIS_SalesOrder 

WHERE STATUS <> 'C' OR STATUS IS NULL and financingcode in (select cuscde from all_customer where acctname = 'ALLIED BANK')


select 
*
--sum(amount) 
from cmis_off_dt where trantype = 'VI' and paidna = 1 and cuscde = 'A00031'


select isnull(sum(baltofinanced),0)-(select isnull(sum(amount),0) from cmis_off_dt where trantype = 'VI' and paidna = 1 and cuscde = 'ALLIED BANK') From SMIS_SalesOrder WHERE STATUS <> 'C' OR STATUS IS NULL and financingcode in (select cuscde from all_customer where acctname = 'ALLIED BANK')



select *,DATEDIFF(day, ar.invoicedate, getdate()) from amis_ar ar 
	inner join amis_detail dt on 
		ar.invoicetype=dt.invoicetype and ar.invoiceno=dt.invoiceno and ar.customercode=dt.customercode and ar.account_code=dt.acct_code
	inner join all_customer_terms tr on ar.customercode=tr.customercode
where DATEDIFF(day, ar.invoicedate, getdate()) > tr.creditterm 
	and ar.amount_topay>dt.invoiceamount
--ar.invoicetype = 'VI' and ar.customercode = 


SELECT *,DATEADD(month,(terms_expiry_year * 12)+terms_expiry_month, getdate()) as expiry FROM ALL_CUSTOMER_TERMS WHERE customercode in (select cuscde from all_customer where acctname = 'ALLIED BANK')


select * from amis_ar ar inner join amis_detail dt on ar.InvoiceType = dt.InvoiceType 
And ar.INVOICENO = dt.INVOICENO And ar.CustomerCode = dt.CustomerCode And ar.account_code = dt.acct_code 
inner join all_customer_terms tr on ar.customercode=tr.customercode 
Where DateDiff(Day, ar.invoicedate, getdate()) > tr.creditterm 
and ar.amount_topay>dt.invoiceamount 
and ar.customercode in (select cuscde from all_customer where acctname='BANCO DE ORO') and ar.status='P'
and ar.invoicetype = 'VI' 