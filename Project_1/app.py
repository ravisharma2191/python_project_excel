import smtplib
import sys
sys.path.append(r"E:\python")
from email.message import EmailMessage
from  DbConfig.db_connectionlive import get_connection
import smtpconfig
import html

# -------------------------------
conn = get_connection()

if not conn:
    print("DB Connection Failed")
    exit()

cursor = conn.cursor()

# -------------------------------
send_email = False
summary = {}

# -------------------------------
queries = {
    "Duplicate Refund Details (Last 15 Days)": {
        "type": "table",
        "sql": """select rr.id, rr.RE_RQE_NO, rr.amount,
                         r.refund_no, r.refund_amount, r.createdby, r.remarks
                  from refund_request rr
                  inner join refund_lineitems rfl on rfl.refund_req_id = rr.id
                  inner join refund r on r.refund_id = rfl.refund_id
                  where rr.id in (
                      select rr.id
                      from refund_request rr
                      inner join refund_lineitems rfl on rfl.refund_req_id = rr.id
                      group by rr.id
                      having count(rfl.refund_req_id) > 1
                  )
                  and trunc(rr.createddatetime) > trunc(sysdate-15)
                  order by rr.id"""
    },

    "Duplicate Invoices": {
        "type": "table",
        "sql": """select org.esiofficename as Unit, P.MRNO,ap.admissionno, p.patientname,
                         v.patient_visit_id,count(im.invoice_no)as billcount,
                         listagg(im.invoice_no,', ') within group(order by im.invoice_no) as Invoiceno
                  from patient p
                  inner join visit v on p.patient_id=v.patient_id
                  inner join patientadmission ap on ap.patientid=v.patient_id and ap.visitid=v.visitid
                  inner join ins_master_invoice im on im.patient_id=v.patient_id and im.visit_id=v.visitid
                  inner join orgstructure org on org.id = im.hospital_id
                  where trunc(im.createddatetime) > to_date('2026-02-01','yyyy-mm-dd')
                  and im.INVOICE_STATUS='Settled'
                  group by org.esiofficename,P.MRNO,ap.admissionno,p.patientname,v.patient_visit_id
                  having count(im.invoice_no)>1"""
    },

    "Receipt PAYCURRENCYID 0": {
        "type": "table",
        "sql": """select ord.esiofficename as unit,PAYCURRENCYID,rc.receiptcode,
                         rc.createddt,rc.receipt_id,rc.receiptamount
                  from payment_details rd
                  inner join receipts rc on rc.receipt_id=rd.receipt_id
                  inner join orgstructure ord on ord.id=rc.siteid
                  where (rd.PAYCURRENCYID='0' or rd.PAYCURRENCYID=14)
                  and rc.createddt>=to_date('01-01-2025','dd-mm-yyyy')
                  order by rc.createddt desc"""
    },

    "User add , in Receipt Deatils": {
        "type": "table",
        "sql": """select ORGS.ID,orgs.esiofficename as unit,r.receiptcode,
                         r.receipt_id,r.receiptdate,pd.ccno,pd.chequeno,pd.utrno
                  from receipts r
                  inner join payment_details pd on pd.receipt_id=r.receipt_id
                  inner join orgstructure orgs on orgs.id=r.siteid
                  AND ORGS.ID NOT IN (17423465)
                  where r.receiptdate >= to_date('01-04-2024','dd-mm-yyyy')
                  and (pd.ccno like'%,%' or pd.chequeno like'%,%'or pd.utrno like'%,%')"""
    },

    "Cancel Days for Invoice >1": {
        "type": "table",
        "sql": """select o.esiofficename as unit,dcd.PARAMETERVALUE canceldays,
                         dcd.UPDATEDDATETIME,u.username
                  from default_configuration dc
                  inner join default_config_details dcd on dcd.config_id=dc.id
                  inner join orgstructure o on o.id=dcd.site_id
                  left join hisuser u on u.id=dcd.UPDATEDBY
                  where o.ISACTIVE=1 and dcd.config_id=555 and dcd.PARAMETERVALUE>1"""
    },

"Cash Refund Amount Limit >5000": {
        "type": "table",
        "sql": """select o.esiofficename as unit,dcd.PARAMETERVALUE CashRefund,dcd.UPDATEDDATETIME,u.username ,dcd.config_id from default_configuration dc
            inner join default_config_details dcd on dcd.config_id=dc.id
            inner join orgstructure o on o.id=dcd.site_id
            left join hisuser u on  u.id=dcd.UPDATEDBY
            where o.ISACTIVE=1 and dcd.config_id=39318032 and dcd.PARAMETERVALUE>5000"""
    },

    "Cash Transaction Limit >199001": {
        "type": "table",
        "sql": """select o.esiofficename as unit,dcd.PARAMETERVALUE Cashtrxlimit,dcd.UPDATEDDATETIME,u.username ,dcd.config_id from default_configuration dc
                    inner join default_config_details dcd on dcd.config_id=dc.id
                    inner join orgstructure o on o.id=dcd.site_id
                    left join hisuser u on  u.id=dcd.UPDATEDBY
                    where o.ISACTIVE=1 and dcd.config_id=494 and dcd.PARAMETERVALUE>199001"""
    },

     "Higher OP Pharmacy Discount >10": {
        "type": "table",
        "sql": """select o.esiofficename as unit,dcd.PARAMETERVALUE OP_Pharmacy_Discount,dcd.UPDATEDDATETIME,u.username ,dcd.config_id from default_configuration dc
                inner join default_config_details dcd on dcd.config_id=dc.id
                inner join orgstructure o on o.id=dcd.site_id
                left join hisuser u on  u.id=dcd.UPDATEDBY
                where o.ISACTIVE=1 and dcd.config_id=604 and dcd.PARAMETERVALUE<>10"""
    },

     "Lower OP Pharmacy Discount >5": {
        "type": "table",
        "sql": """select o.esiofficename as unit,dcd.PARAMETERVALUE CashRefund,dcd.UPDATEDDATETIME,u.username ,dcd.config_id from default_configuration dc
                inner join default_config_details dcd on dcd.config_id=dc.id
                inner join orgstructure o on o.id=dcd.site_id
                left join hisuser u on  u.id=dcd.UPDATEDBY
                where o.ISACTIVE=1 and dcd.config_id=603 and dcd.PARAMETERVALUE<>5 and o.id<>17423465"""
    },
    
    "Service Master Point(.) in Rate": {
        "type": "table",
        "sql": """select o.esiofficename as unit,t.id,t.service_id,
                         sm.service_code,sm.service_name,t.name,
                         t.totalcharges,tc.TARIFFGROUP_NAME as TeariffName,
                         case when tg.TARIFFGROUP_NAME is null 
                              then tc.TARIFFGROUP_NAME 
                              else tg.TARIFFGROUP_NAME end Tariffclass,
                         t.createddatetime
                  from tariff t
                  left join tariffgroup tg on tg.id=t.tariffgroupid
                  inner join orgstructure o on o.id=t.hospitalid
                  left join tariffgroup tc on tc.id=t.tarffic_class_id
                  inner join servicemaster sm on sm.service_master_id=t.service_id
                  where t.totalcharges like'%.%' and t.active=1"""
    },
    "Invoice not Transfer to SAP": {
        "type": "table",
        "sql": """select inv.hospital_id,inv.createddatetime,inv.updateddatetime,inv.invoice_no,inv.invoice_status,inv.docnum_sap 
,inv.NET_AMOUNT,sum(il.discount_amt)
from ins_master_invoice inv
inner join ins_invoice_lineitems il on il.master_inv_id=inv.master_invoice_id
where  trunc(inv.createddatetime) BETWEEN trunc(Sysdate-30)
        and trunc(Sysdate-1) and inv.ISREVENUESTAGEPOSTED=0 and inv.INVOICE_STATUS='Settled'
        and inv.APPROVALSTATUS=2 and inv.hospital_id not in ('17423465','17423471','17423468','17423469') 
        and inv.NET_AMOUNT>0 and trunc(inv.updateddatetime)<=(Sysdate-1)
       group by inv.hospital_id,inv.createddatetime,inv.updateddatetime,
       inv.invoice_no,inv.invoice_status,inv.docnum_sap,inv.NET_AMOUNT
        having inv.NET_AMOUNT<>sum(il.discount_amt)
        """
    },
    "package location not mapped": {
        "type": "table",
        "sql": """select pd.packagecode,pd.packagename,pd.createddatetime,u.username,pd.updateddatetime
                    from packagedefinition pd
                    inner join hisuser u on u.id=pd.createdby
                    where pd.servicemasterid not in (
                    SELECT service_id FROM servicelocationmap where service_id IS NOT NULL)
                    and active=1  --and  trunc(pd.createddatetime)>to_date('01-01-2019','dd-mm-yyyy')
                    and pd.servicemasterid not in(select service_id from ins_invoice_lineitems)
                    and pd.packagecode not like'%Health%'
                    order by u.username,pd.createddatetime
        """
    },
    "Service location not mapped": {
        "type": "table",
        "sql": """select sm.service_master_id ,sm.service_code,sm.service_name,sm.createddatetime,u.username,sm.updateddatetime
                    from servicemaster sm
                    inner join hisuser u on u.id=sm.createdby
                    where sm.service_master_id not in (
                    SELECT service_id FROM servicelocationmap where service_id IS NOT NULL)
                    and sm.is_active='Y'  and  trunc(sm.createddatetime)>to_date('01-01-2024','dd-mm-yyyy')
                    and sm.service_master_id not in(select servicemasterid from packagedefinition)
                     --and sm.service_master_id  not in (select serviceid from directservicetariff)
                     --and sm.service_master_id  in (select service_id from tariff) 
                    order by u.username,sm.createddatetime"""
    }

    ,
    "Service not chargable": {
        "type": "table",
        "sql": """select * from servicemaster where chargable='N' and is_active='Y'
                    and trunc(createddatetime) > to_date('01-01-2023','dd-mm-yyyy')
        """
    },
    "Package Tariff 0 Rate": {
        "type": "table",
        "sql": """select pd.packagecode,pd.servicemasterid,t.id,t.totalcharges from packagedefinition pd
            inner join tariff t on t.service_id=pd.servicemasterid and t.active=1
            and t.totalcharges=0 and TARIFFGROUPID<>1742381 
            where pd.active=1 and trunc(pd.createddatetime)> to_date('01-01-2023','dd-mm-yyyy')"""
    },
     "GRN Stock Not Update": {
        "type": "table",
        "sql": """select grnno,grndate,ORGSTRUCTCODE,orgs.esiofficename,s.description from goodsreceiptnote 
            inner join orgstructure orgs on ORGSTRUCTCODE=orgs.id
            inner join store s on s.id=goodsreceiptnote.purchasestore
            where grnno not in (select refdocnum from stockledger)
            and trunc(grndate)= trunc(sysdate-1) and APPROVESTATUSTYPENUM='2' and GRNTYPE<>1"""
    }
}

# -------------------------------
text_body = "Daily Activity Report\n\n"
html_body = """
<html>
<body>
<h2>Daily Activity Report</h2>
"""

# -------------------------------
for name, q in queries.items():
    try:
        cursor.execute(q["sql"])

        if q["type"] == "table":
            rows = cursor.fetchall()
            count = len(rows)

            summary[name] = count

            if count > 0:
                send_email = True

                columns = [col[0] for col in cursor.description]

                html_body += f"<h3>{name}</h3>"
                html_body += "<table border='1'>"

                html_body += "<tr>"
                for col in columns:
                    html_body += f"<th>{col}</th>"
                html_body += "</tr>"

                for row in rows:
                    html_body += "<tr>"
                    for cell in row:
                        html_body += f"<td>{html.escape(str(cell))}</td>"
                    html_body += "</tr>"

                html_body += "</table><br>"

    except Exception as e:
        summary[name] = "Error"
        html_body += f"<p>Error in {name}: {e}</p>"

cursor.close()
conn.close()

# -------------------------------
# SUMMARY SECTION
# -------------------------------
summary_html = """
<h2>Summary Report</h2>
<table border='1'>
<tr><th>Report Name</th><th>Count</th></tr>
"""

for k, v in summary.items():
    summary_html += f"<tr><td>{k}</td><td>{v}</td></tr>"

summary_html += "</table><br>"

# Insert summary at top
html_body = html_body.replace(
    "<h2>Daily Activity Report</h2>",
    "<h2>Daily Activity Report</h2>" + summary_html
)

#html_body += "</body></html>"

# -------------------------------
if not send_email:
    html_body += "<p><b>No data found for any query.</b></p>"

html_body += "</body></html>"

text_body = "Daily Activity Report - Summary Attached"

# -------------------------------
msg = EmailMessage()
msg["Subject"] = "Daily Activity Report"
msg["From"] = smtpconfig.SENDER_EMAIL
msg["To"] = ", ".join(smtpconfig.RECEIVER_EMAILS)

msg.set_content(text_body)
msg.add_alternative(html_body, subtype="html")

# -------------------------------
try:
    with smtplib.SMTP(smtpconfig.SMTP_SERVER, smtpconfig.SMTP_PORT) as server:
        server.starttls()
        server.login(smtpconfig.SENDER_EMAIL, smtpconfig.SENDER_PASSWORD)
        server.send_message(msg)

    print("Mail sent successfully")

except Exception as e:
    print("Mail Error:", e)