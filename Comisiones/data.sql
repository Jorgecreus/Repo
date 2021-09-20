with warranty as
(select
distinct

o.order_number as order_number,
(CASE WHEN roi.pnl_cost_price_gross = '26620' AND vat_key='R' THEN 'ES - Autohero Plata - 24'
        WHEN roi.pnl_cost_price_gross = '43560' AND vat_key='R' THEN 'ES - Autohero Plata - 36'
        WHEN roi.pnl_cost_price_gross = '14520' AND vat_key='R' THEN 'ES - Autohero Oro - 12'
        WHEN roi.pnl_cost_price_gross = '31944' AND vat_key='R' THEN 'ES - Autohero Oro - 24'
        WHEN roi.pnl_cost_price_gross = '52272' AND vat_key='R' THEN 'ES - Autohero Oro - 36'
        WHEN roi.pnl_cost_price_gross = '27830' AND vat_key='R' THEN 'ES - Autohero Diamante - 12'
        WHEN roi.pnl_cost_price_gross = '52514' AND vat_key='R' THEN 'ES - Autohero Diamante - 24'
        WHEN roi.pnl_cost_price_gross = '81312' AND vat_key='R' THEN 'ES - Autohero Diamante - 36'
        ELSE 'ES - Autohero Plata - 12'
        END) AS warranty_title,
o.car_handover_on        
        
from wkda_retail.retail_order as o
LEFT JOIN wkda_retail.retail_order_item as roi ON  roi.order_id = o.id

where roi.vat_key='R'
AND roi.price_gross in ('0', '29900', '54900', '24900', '39900', '64900', '39900', '59900', '94900')
AND o.car_handover_on >= '2021-08-01 00:00:00' 
--and o.car_handover_on <= '2021-08-31 23:59:59'
--and o.canceled_on is null

group by 1,2,3)

select
distinct
--Col1
opps.stock_number,
--Col2
--o.order_number,
case when h.order_number is null then o1.order_number else h.order_number end as order_number,
--Col3
us.email_address,
--Col4
h.booking_date,
--Col5
--opd.payment_type,
case when h.payment_type is null then opd.payment_type else h.payment_type end as payment_type,
--Col6
date(opps.contract_signed_on) as contract_signed_on,
--Col7
w.warranty_title,
--Col8
date(opps.car_handover_on) as car_handover_on,
--Col9
nps.nps_value

--Col10
--h.order_id,
--Col11
--max(opps.first_interaction_created_on) as first_interaction_created_on,
--Col12
--opps.conversion_type,
--Col13
--(opps.assigned_to_firstname|| ' ' ||opps.assigned_to_lastname) as name
--Col14
--h.order_created_by,
--Col15
--opps.is_lead_archived
--Col16
--max(roi.pnl_cost_price_gross),
--Col17
--max(roi.updated)


from ba_kr_retail_clm as opps
left join ah_eds_hero_table as h on h.stock_number=opps.stock_number
left join wkda_clm.opportunity_property as op on op.value=h.order_number
left join dwh_load.nps_temp_raw_data as nps on nps.order_number=h.order_number
left join wkda_retail.retail_order as o on o.stock_number=opps.stock_number
left join wkda_retail.retail_order as o1 on o1.id=opps.order_id
left join wkda_retail.retail_order_payment_details as opd on opd.order_id=o1.id
left join wkda_clm.opportunity as opor on opor.uuid=opps.opp_id
left join wkda_clm_user_info.user_info as us on opps.assigned_to = us.uuid
--LEFT JOIN wkda_retail.retail_order_item as roi ON  roi.order_id = o.id
left join warranty as w on w.order_number=o1.order_number

where opps.retail_country='ES'
--and h.order_status='COMPLETED'
and opps.state_ra='DELIVERED_TO_CUSTOMER'
--and o.state='COMPLETED'
--and o.state='DELIVERED'
--and opps.state_ro='COMPLETED'
--and opps.state_ro='DELIVERED'
--and op.name='ORDER_NUMBER'
and opps.contract_signed_on != 0
and opps.car_handover_on != 0
--and opor.status='SUCCESSFUL'
--and opps.is_lead_archived !=1
and us.email_address !=''
--and w.warranty_title !=0
and opps.car_handover_on >= '2021-08-01 00:00:00' 
--and opps.car_handover_on <= '2021-08-31 23:59:59'

--and opps.stock_number in ('DH23764', 'VH78283')

Group by 1,2,3,4,5,6,7,8,9
order by 1 desc
