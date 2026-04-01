select id,employee_id,pool_id,batch_id,grant_period_id,no_vest,stock_type_id,stock_qty,exercise_list_id,exercised_qty
from esop_grant order by id desc limit 100;