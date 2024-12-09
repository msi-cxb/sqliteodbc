select 'select
' || group_concat('    case when [' || name || '] is not null then ' || quote(name || ', ') || ' else '''' end', ' ||
') || '
  as columns,
  count(*) as num_rows
from
  [' || '_T_' || ']
group by
  columns
order by
  num_rows desc' as query from pragma_table_info('_T_')
