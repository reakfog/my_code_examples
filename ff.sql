-- 1. Procurement korrection correction --

-- таблица с данными из weekly
with weekly as(
    select weekly.id_product,
           weekly.rpr,
         --weekly.total_pv_ab as pv_ab,
           psb.psb,
           weekly.pwo_total as pwo
    from data._weekly_dataset(current_date - interval '7 days 3 hours', current_date - interval '3 hours') as weekly
    left join data.da_psb(current_date - interval '7 days 3 hours',current_date - interval '3 hours') psb on weekly.id_product = psb.id_product
),
-- таблица с процентом поставок за неделю
slp as(
    select slp.id_product,
           sum(slp.fact) / sum(case when slp.purchase_nominal is null or slp.package_count is null then slp.count_original
                                    else slp.purchase_nominal * slp.package_count
                                    end) * 100 as supply
    from data.da_slp(current_date - interval '7 days 3 hours', current_date - interval '3 hours') as slp
    group by slp.id_product
),
-- таблица сроков годности
shelflife as(
    select tproductstorageconditions.id_product,
           case when tproductstorageconditions.id_shelflifewunit = 8 then 365 * tproductstorageconditions.shelflife
                when tproductstorageconditions.id_shelflifewunit = 9 then (365::numeric/12::numeric) * tproductstorageconditions.shelflife
                when tproductstorageconditions.id_shelflifewunit = 10 then 1 * tproductstorageconditions.shelflife
                when tproductstorageconditions.id_shelflifewunit = 11 then (1::numeric/24::numeric) * tproductstorageconditions.shelflife
                when tproductstorageconditions.id_shelflifewunit = 12 then 1 * tproductstorageconditions.shelflife
                else 9999
                end as shelf_life
    from tproductstorageconditions
),
-- вывод таблицы для расчетов
result as(
    select weekly.id_product,
           coalesce(slp.supply, 100) as supply,
           weekly.rpr,
           weekly.psb,
           weekly.pwo,
           shelflife.shelf_life
    from weekly
    left join slp on slp.id_product = weekly.id_product
    left join shelflife on shelflife.id_product = weekly.id_product
),
-- настройка расчета коэфициента в зависимости от psb
result_2 as(
    select result.*,
           case when result.supply = 0 or result.psb = 0 then 100
                else 100 * result.psb / result.supply
                end as estimated_pv_ab
    from result
),
-- настройка расчета коэфициента в зависимости от pwo
result_3 as(
    select result_2.*,
           case when result_2.pwo <= 10 or (result_2.pwo <= 15 and result_2.rpr >= 200) then 100 * 1 / result_2.estimated_pv_ab
                else 1
                end as k_element
    from result_2
),
-- настройка расчета коэфициента в зависимости от срока годности
result_4 as(
    select result_3.*,
           case when result_3.shelf_life * 0.8 < 3
                then (case when result_3.k_element > 2 then 2
                           else result_3.k_element
                           end)
                when result_3.shelf_life * 0.8 < 15
                then (case when result_3.k_element > 2.5 then 2.5
                           else result_3.k_element
                           end)
                when result_3.shelf_life * 0.8 < 30
                then (case when result_3.k_element > 3 then 3
                           else result_3.k_element
                           end)
                else (case when result_3.k_element > 3.5 then 3.5
                           else result_3.k_element
                           end)
                end as k
    from result_3
)

select result_4.id_product,
       result_4.k,
       concat('update tdistributionproviderproducts set purchase_correctioncoef = ',result_4.k,' where id_product = ',result_4.id_product,';') as request
from result_4;



-- 2. PSB --
create function da_psb(
    pdatefrom timestamp without time zone DEFAULT NULL::timestamp without time zone,
    pdateto timestamp without time zone DEFAULT NULL::timestamp without time zone)
    returns TABLE(id_product bigint, name_product character varying, psb double precision)
    language plpgsql
as
$$
begin
    RETURN query

with table_unit_prod_dc as(
    with base_table as(
        select tproducts.id as id_product,
               tproducts.name as name_product,
               tproductitems.id_storepoint,
               tunits.id as id_wunit,
               tunits.name as unit_prod,
               tdimensiontypes.id as id_dimensiontype,
               tdimensiontypes.name,
               case when tdimensiontypes.id in (2, 3) then 1
                    when tdimensiontypes.id = 1 and tunits.id = 2 then tproductitems.weight / 1000.0
                    when tdimensiontypes.id = 1 and tunits.id = 1 then tproductitems.weight
                    end as unit_prod_count
        from tproducts
        inner join tproductitems tproductitems on tproducts.id = tproductitems.id_product
        inner join tunits on tproductitems.id_wunit = tunits.id
        inner join tdimensiontypes on tproductitems.dimensiontype = tdimensiontypes.id
        where tproductitems.id_storepoint in (select tstorepoints.id from tstorepoints where tstorepoints.isactive is true)
    )

    select base_table.id_product,
           base_table.id_storepoint, 
           coalesce(base_table.unit_prod_count, 1) as unit_prod_count
    from base_table
),

table_unit_prod_hub as(
    select distinct tproducts.id, tproducts.name,
           case when stock_dimensiontype in (2, 3) then 1
                when stock_dimensiontype = 1 and id_wunit = 2 then weight / 1000
                when stock_dimensiontype = 1 and id_wunit = 1 then weight end as unit_prod_count
    from tdistributionproviderproducts
    inner join tproducts on tdistributionproviderproducts.id_product = tproducts.id
    where tdistributionproviderproducts.is_actual is true
),

table_dc_instock as(
    select distinct tproducts.id,
           tproducts.name,
           tproductitems.id_storepoint,
           tstockhistory.date,
           tstockhistory.instock_count,
           tstockhistory.id_type,
           tstockhistory.id as ids,
           coalesce(table_unit_prod_dc.unit_prod_count, 1) as unit_prod_count,
           coalesce(lag(tstockhistory.date, -1) over (partition by tproducts.id, tproductitems.id_storepoint order by tstockhistory.id),  pdateto :: timestamp) as lg
    from tstockhistory
    inner join tproductitems on tstockhistory.id_productitem = tproductitems.id
    inner join tproducts on tproductitems.id_product = tproducts.id
    left join table_unit_prod_dc on table_unit_prod_dc.id_product = tproductitems.id_product and table_unit_prod_dc.id_storepoint = tproductitems.id_storepoint
),

table_hub_instock as(
    select tdistributionhistory.id as id_supply,
           tproducts.id as id_product,
           tproducts.name,
           tdistributionhistory.date as date,
           tdistributionhistory.id_type,
           tdistributionhistory.instock_count as provider_instock_count,
           coalesce(lag(tdistributionhistory.date, -1) over (partition by tproducts.id order by tdistributionhistory.id), pdateto :: timestamp) as lg,
           coalesce(table_unit_prod_hub.unit_prod_count, 1) as unit_prod_count
    from tdistributionhistory
    inner join tproviderproducts on tdistributionhistory.id_providerproduct = tproviderproducts.id
    inner join tproducts on tproviderproducts.id_product = tproducts.id
    left join table_unit_prod_hub on table_unit_prod_hub.id = tproducts.id
),

hub_filter_instock as(
    select distinct table_hub_instock.id_product,
           table_hub_instock.name as name_product,
           table_hub_instock.date,
           table_hub_instock.provider_instock_count as provider_instock_count,
           table_hub_instock.unit_prod_count,
           case when lg > pdateto then pdateto
                else lg
                end as lg
    from table_hub_instock
    where table_hub_instock.provider_instock_count < table_hub_instock.unit_prod_count and
          table_hub_instock.lg >= pdatefrom and
          table_hub_instock.date < pdateto
),

dc_filter_instock as(
    select table_dc_instock.id,
           table_dc_instock.name,
           table_dc_instock.id_storepoint,
           table_dc_instock.date,
           table_dc_instock.instock_count,
           case when lg > pdateto then pdateto
                else lg
                end as lg,
           unit_prod_count
    from table_dc_instock
    where table_dc_instock.instock_count < unit_prod_count and table_dc_instock.lg >= pdatefrom
    and table_dc_instock.date < pdateto
),

instock_all as(
    with table_last_dc as(
        with tb1 as(
            select base.id, base.name,
                   base.id_storepoint,
                   max(base.rn) as mx
            from(
                select distinct tproducts.id,
                       tproducts.name,
                       tproductitems.id_storepoint,
                       tstockhistory.date,
                       tstockhistory.instock_count,
                       tstockhistory.id_type, 
                       tstockhistory.id as ids,
                       row_number() over (partition by tproducts.id, tproductitems.id_storepoint order by tstockhistory.id) as rn
                from tstockhistory
                inner join tproductitems on tstockhistory.id_productitem = tproductitems.id
                inner join tproducts on tproductitems.id_product = tproducts.id
                where tstockhistory.date < pdateto
            ) as base
            group by base.id, base.name, base.id_storepoint
        ),

        tb2 as(
            select distinct tproducts.id, 
                   tproducts.name,
                   tproductitems.id_storepoint,
                   tstockhistory.date,
                   tstockhistory.instock_count,
                   tstockhistory.id_type,
                   tstockhistory.id as ids,
                   row_number() over (partition by tproducts.id, tproductitems.id_storepoint order by tstockhistory.id) as rn
            from tstockhistory
            inner join tproductitems on tstockhistory.id_productitem = tproductitems.id
            inner join tproducts on tproductitems.id_product = tproducts.id
        ),

        tb3 as(
            select tb1.id as id_product,
                   tb1.name as name_product,
                   tb1.id_storepoint,
                   tb2.instock_count,
                   tb2.date
           from tb1
           left join tb2 on tb1.id = tb2.id and tb1.id_storepoint = tb2.id_storepoint and
                     tb1.mx = tb2.rn
        )

        select * from tb3
    ),

    table_last_hub as(
        with tb1 as(
            select base.id,
                   max(base.rn) as mx
            from(
                select tdistributionhistory.id as id_supply,
                       tproducts.id,
                       tdistributionhistory.date as date,
                       tdistributionhistory.id_type,
                       tdistributionhistory.instock_count as provider_instock_count,
                       row_number() over (partition by tproducts.id order by tdistributionhistory.id) as rn
                from tdistributionhistory
                inner join tproviderproducts on tdistributionhistory.id_providerproduct = tproviderproducts.id
                inner join tproducts on tproviderproducts.id_product = tproducts.id
                where tdistributionhistory.date < pdateto
           ) as base
           group by base.id),

        tb2 as(
            select tdistributionhistory.id as id_supply,
                   tproducts.id,
                   tdistributionhistory.date as date,
                   tdistributionhistory.id_type,
                   tdistributionhistory.instock_count as provider_instock_count,
                   row_number() over (partition by tproducts.id order by tdistributionhistory.id) as rn
            from tdistributionhistory
            inner join tproviderproducts on tdistributionhistory.id_providerproduct = tproviderproducts.id
            inner join tproducts on tproviderproducts.id_product = tproducts.id
        ),

        tb3 as (
            select tb1.id as id_product,
                   tb2.date,
                   tb2.provider_instock_count
            from tb1
            left join tb2 on tb1.id = tb2.id and
                      tb1.mx = tb2.rn
        )

        select * from tb3
    ),

    table_result_last as(
        select table_last_dc.id_product,
               table_last_dc.id_storepoint,
               table_last_dc.instock_count,
               table_last_dc.date as date_dc_last,
               table_last_hub.date as date_hub_last,
               table_last_hub.provider_instock_count as provider_instock_count
        from table_last_dc
        left join table_last_hub on table_last_dc.id_product = table_last_hub.id_product
    )

    select * from table_result_last
),

all_get_products as(
    select distinct tproducts.id as id_product,
           tproducts.name as name_product,
           tstorepoints.id as id_storepoint
    from tproducts
    inner join tproductitems on tproducts.id = tproductitems.id_product
    inner join tproviderproducts on tproducts.id = tproviderproducts.id_product
    inner join tstorepoints on tproductitems.id_storepoint = tstorepoints.id
    where tproducts.deleted is not true and
          tproducts.id_internal_type is null and
          tproductitems.isactive is true and
          tproductitems.isinsellout is not true and
          tproviderproducts.isactive is true and
          tstorepoints.isactive is true
),

table_merge as(
    select all_get_products.id_product as id,
           all_get_products.name_product as name,
           all_get_products.id_storepoint,
           dc_filter_instock.date as date_start_instock_dc,
           dc_filter_instock.lg as date_end_instock_dc,
           coalesce(instock_all.date_dc_last, pdatefrom :: timestamp) as date_dc_last,
           coalesce(instock_all.instock_count, 0) as instock_dc_last,
           coalesce(instock_all.date_hub_last, pdatefrom :: timestamp) as date_hub_last,
           coalesce(instock_all.provider_instock_count, 0) as provider_instock_count_last,
           hub_filter_instock.date as date_start_hub,
           hub_filter_instock.lg as date_end_hub,
           pdatefrom :: timestamp as datefrom,
           pdateto :: timestamp as dateto,
           coalesce(table_unit_prod_dc.unit_prod_count, 1) as unit_prod_count_dc,
           coalesce(table_unit_prod_hub.unit_prod_count, 1) as table_unit_prod_hub
    from all_get_products
    left join dc_filter_instock on all_get_products.id_product = dc_filter_instock.id and
              all_get_products.id_storepoint = dc_filter_instock.id_storepoint
    full join hub_filter_instock on hub_filter_instock.id_product = all_get_products.id_product
    left join instock_all on instock_all.id_product = all_get_products.id_product and
              instock_all.id_storepoint = all_get_products.id_storepoint
    left join table_unit_prod_dc on table_unit_prod_dc.id_product = all_get_products.id_product and
              table_unit_prod_dc.id_storepoint = all_get_products.id_storepoint
    left join table_unit_prod_hub on table_unit_prod_hub.id = all_get_products.id_product
),

table_merge_filter as(
    select table_merge.id,
           table_merge.name,
           table_merge.id_storepoint,
           case when table_merge.date_hub_last <= table_merge.datefrom and
                     table_merge.provider_instock_count_last < table_merge.table_unit_prod_hub and
                     table_merge.date_start_hub is null
                then table_merge.datefrom
                else table_merge.date_start_hub
                end as date_start_hub_,
           case when table_merge.date_hub_last <= table_merge.datefrom and
                     table_merge.provider_instock_count_last < table_merge.table_unit_prod_hub and
                     table_merge.date_end_hub is null
                then table_merge.dateto
                else table_merge.date_end_hub
                end as date_end_hub_,
           case when table_merge.date_dc_last <= table_merge.datefrom and
                     table_merge.instock_dc_last < table_merge.unit_prod_count_dc and
                     table_merge.date_start_instock_dc is null
                then table_merge.datefrom
                else table_merge.date_start_instock_dc end as date_start_instock_dc_,
           case when table_merge.date_dc_last <= table_merge.datefrom and
                     table_merge.instock_dc_last < table_merge.unit_prod_count_dc and
                     table_merge.date_end_instock_dc is null
                then table_merge.dateto
                else table_merge.date_end_instock_dc
                end as date_end_instock_dc_,
           table_merge.unit_prod_count_dc,
           table_merge.table_unit_prod_hub,
           table_merge.datefrom,
           table_merge.dateto,
           table_merge.date_hub_last,
           table_merge.provider_instock_count_last,
           table_merge.date_start_hub,
           table_merge.date_end_hub,
           table_merge.date_dc_last,
           table_merge.instock_dc_last,
           table_merge.date_start_instock_dc,
           table_merge.date_end_instock_dc
    from table_merge
),

table_merge_clear as(
    select table_merge_filter.id as id_product,
           table_merge_filter.name,
           table_merge_filter.id_storepoint,
           case when table_merge_filter.date_start_hub_ < pdatefrom then pdatefrom :: timestamp
                else table_merge_filter.date_start_hub_
                end as date_start_hub_,
           case when table_merge_filter.date_end_hub_ > pdateto then pdateto :: timestamp
                else table_merge_filter.date_end_hub_
                end as date_end_hub_,
           case when table_merge_filter.date_start_instock_dc_ < pdatefrom then pdatefrom :: timestamp
                else table_merge_filter.date_start_instock_dc_
                end as date_start_instock_dc_,
           case when table_merge_filter.date_end_instock_dc_ > pdateto then pdateto :: timestamp
                else table_merge_filter.date_end_instock_dc_
                end as date_end_instock_dc_,
           table_merge_filter.unit_prod_count_dc,
           table_merge_filter.table_unit_prod_hub,
           table_merge_filter.datefrom,
           table_merge_filter.dateto,
           table_merge_filter.date_hub_last,
           table_merge_filter.provider_instock_count_last,
           table_merge_filter.date_start_hub,
           table_merge_filter.date_end_hub,
           table_merge_filter.date_dc_last,
           table_merge_filter.instock_dc_last,
           table_merge_filter.date_start_instock_dc,
           table_merge_filter.date_end_instock_dc
    from table_merge_filter
),

table_tsrange as(
    select table_merge_clear.id_product,
           table_merge_clear.name,
           table_merge_clear.id_storepoint,
           tsrange(table_merge_clear.date_start_hub_, table_merge_clear.date_end_hub_) &&
           tsrange(table_merge_clear.date_start_instock_dc_, table_merge_clear.date_end_instock_dc_) and
                  table_merge_clear.date_start_hub_ is not null as ts,
           table_merge_clear.unit_prod_count_dc,
           table_merge_clear.table_unit_prod_hub,
           table_merge_clear.date_hub_last,
           table_merge_clear.date_start_hub,
           table_merge_clear.date_end_hub,
           table_merge_clear.date_start_instock_dc,
           table_merge_clear.date_end_instock_dc,
           table_merge_clear.date_start_hub_,
           table_merge_clear.date_end_hub_,
           table_merge_clear.date_start_instock_dc_,
           table_merge_clear.date_end_instock_dc_
    from table_merge_clear
),

table_tsrange_filter as(
    select table_tsrange.id_product as id_product,
           table_tsrange.name as name_product,
           table_tsrange.id_storepoint,
           case when table_tsrange.date_start_hub_ >= table_tsrange.date_start_instock_dc_ and
                     table_tsrange.date_end_hub_ >= table_tsrange.date_end_instock_dc_ and
                     table_tsrange.ts is true
                then extract(epoch from table_tsrange.date_end_instock_dc_ - table_tsrange.date_start_hub_)
                when table_tsrange.date_start_hub_ <= table_tsrange.date_start_instock_dc_ and
                     table_tsrange.date_end_instock_dc_ <= table_tsrange.date_end_hub_ and
                     table_tsrange.ts is true
                then extract(epoch from table_tsrange.date_end_instock_dc_ - table_tsrange.date_start_instock_dc_)
                when table_tsrange.date_start_instock_dc_ >= table_tsrange.date_start_hub_ and
                     table_tsrange.date_end_instock_dc_ >= table_tsrange.date_end_hub_ and
                     table_tsrange.ts is true
                then extract(epoch from table_tsrange.date_end_hub_ - table_tsrange.date_start_instock_dc_)
                when table_tsrange.date_start_hub_ >= table_tsrange.date_start_instock_dc_ and
                     table_tsrange.date_end_hub_ <= table_tsrange.date_end_instock_dc_ and
                     table_tsrange.ts is true
                then extract(epoch from table_tsrange.date_end_hub_ - table_tsrange.date_start_hub_)
                end as sec_,
           table_tsrange.ts
    from table_tsrange
    where table_tsrange.ts is true
),

table_tsrange_sum as(
    select table_tsrange_filter.id_product,
           table_tsrange_filter.name_product,
           table_tsrange_filter.id_storepoint,
           coalesce(sum(table_tsrange_filter.sec_), 0) as sec_
    from table_tsrange_filter
    group by table_tsrange_filter.id_product, table_tsrange_filter.name_product, table_tsrange_filter.id_storepoint
),

table_result as(
    select table_tsrange_sum.id_product,
           table_tsrange_sum.name_product,
           table_tsrange_sum.id_storepoint,
           table_tsrange_sum.sec_,
           (count(*) over (partition by table_tsrange_sum.id_product)) as count_string,
           ((1 - table_tsrange_sum.sec_ / (extract (epoch from (pdateto :: timestamp - pdatefrom :: timestamp)))) * (1 / (select count(*) from tstorepoints where tstorepoints.isactive is true) :: double precision)) :: double precision as res
    from table_tsrange_sum
),

table_psb as(
    select all_get_products.id_product,
           all_get_products.name_product,
           coalesce((sum(table_result.res) + (((select count(*) from tstorepoints where isactive is true) - avg(count_string)) * (1 / (select count(*) from tstorepoints where tstorepoints.isactive is true) :: double precision))) * 100, 100) :: double precision as result
    from all_get_products
    left join table_result on table_result.id_product = all_get_products.id_product and
         table_result.id_storepoint = all_get_products.id_storepoint
    group by all_get_products.id_product, all_get_products.name_product
)

select table_psb.id_product :: bigint as id_product,
       table_psb.name_product :: varchar as name_product,
       table_psb.result :: double precision as psb
from table_psb
order by table_psb.id_product;

end
$$;

alter function da_psb(timestamp, timestamp) owner to sophiedata;



-- 3. PV
create function pv_function(
    pdatefrom timestamp without time zone DEFAULT NULL::timestamp without time zone,
    pdateto timestamp without time zone DEFAULT NULL::timestamp without time zone,
    pid_products integer[] DEFAULT NULL::integer[])
    returns TABLE(id_product bigint, name_product character varying, pv_ab double precision, pv double precision)
    language plpgsql
as
$$
begin
  RETURN query

with all_get_products as(
    select distinct tproducts.id as id_product,
           tproducts.name as name_product,
           tproductitems.id_storepoint
    from tproducts
    join tproductitems on tproducts.id = tproductitems.id_product
    join tproviderproducts on tproducts.id = tproviderproducts.id_product
    join tstorepoints on tproductitems.id_storepoint = tstorepoints.id
    where tproducts.deleted is not true and
          tproducts.id_internal_type is null and
          tproductitems.isactive is true and
          tproductitems.isinsellout is not true and
          tproviderproducts.isactive is true and
          tstorepoints.isactive is true
    order by id_product
),

all_times as(
    select tproducts.id,
           tproducts.name,
           tproductitems.id_storepoint,
           tproductitemsalehistory.date_start,
           tproductitemsalehistory.date_end,
           case when tproductitemsalehistory.date_start < pdatefrom and
                     ((tproductitemsalehistory.date_end > pdatefrom and tproductitemsalehistory.date_end < pdateto) or
                     (tproductitemsalehistory.date_end is null and tproductitemsalehistory.date_start < pdateto) or
                     tproductitemsalehistory.date_end > pdateto)
                then make_timestamp(extract(year from pdatefrom :: timestamp) :: int, extract(month from pdatefrom :: timestamp) :: int, extract(day from pdatefrom :: timestamp) :: int, extract(hour from pdatefrom :: timestamp) :: int, 0, 0)
                else tproductitemsalehistory.date_start
                end as new_date_start,
           case when (tproductitemsalehistory.date_end > pdateto and tproductitemsalehistory.date_start < pdateto) or (tproductitemsalehistory.date_end is null and tproductitemsalehistory.date_start < pdateto)
                then make_timestamp(extract(year from pdateto :: timestamp) :: int, extract(month from pdateto :: timestamp) :: int, extract(day from pdateto :: timestamp) :: int, extract(hour from pdateto :: timestamp) :: int, 0, 0)
                else tproductitemsalehistory.date_end
                end as new_date_end
    from tproductitemsalehistory
    left join tproductitems on tproductitems.id = tproductitemsalehistory.id_productitem
    left join tproducts on tproducts.id = tproductitems.id_product
    join tstorepoints on tstorepoints.id = tproductitems.id_storepoint
    join tproviders on tproviders.id = tproductitems.id_provider
    where tproducts.id in (select id from tproducts) and
          tproducts.id_internal_type is null and
          tproducts.deleted is not true and
          tproductitems.isinsellout is not true and
          tproductitems.id_storepoint in (select distinct tstorepoints.id from tstorepoints where tstorepoints.isactive is true) and
          tproductitems.isactive is true and
          tstorepoints.isactive is true and
          tproviders.isactive is true
),

get_value as(
    select all_times.id,
           all_times.name,
           all_times.id_storepoint,
           all_times.date_start,
           all_times.date_end,
           all_times.new_date_start,
           all_times.new_date_end,
           extract(epoch from all_times.new_date_end - all_times.new_date_start) as value_diff
    from all_times
    where (pdatefrom IS NULL or new_date_start >= pdatefrom) and
          (pdateto IS NULL or new_date_end <= pdateto) and
          (pdateto IS NULL or new_date_end <= pdateto)
),

get_result as(
    select all_get_products.id_product as id,
           all_get_products.name_product as name,
           all_get_products.id_storepoint,
           coalesce(sum(get_value.value_diff), 0) as value_diff
    from all_get_products
    left join get_value on get_value.id = all_get_products.id_product and
              get_value.id_storepoint = all_get_products.id_storepoint
    group by all_get_products.id_product, all_get_products.name_product, all_get_products.id_storepoint
),

pv as(
    select get_result.id,
           get_result.name,
           get_result.id_storepoint,
           get_result.value_diff / extract(epoch from (pdateto :: timestamp - pdatefrom :: timestamp)) * 100 as pv
    from get_result
),

table_pv_ab as(
    with table_all_products as(
        select tproducts.id as id_product,
               tproducts.name as name_product,
               tproductitems.id_storepoint,
               round(count(*) filter ( where ab is true ) * 15.0 * 60.0 :: double precision / extract(epoch from (pdateto :: timestamp - pdatefrom :: timestamp)) :: double precision * 100, 2) as pv_ab,
               round(count(*) filter ( where instant is true ) * 15.0 * 60.0 :: double precision / extract(epoch from (pdateto :: timestamp - pdatefrom :: timestamp)) :: double precision * 100, 2) as pv_storepoint
        from  analytics_product_oos_byproduct
        join tproductitems on analytics_product_oos_byproduct.id_productitem = tproductitems.id
        join tproducts on tproductitems.id_product = tproducts.id
        where analytics_product_oos_byproduct.date >= pdatefrom and
              analytics_product_oos_byproduct.date < pdateto
        group by tproducts.id, tproducts.name, tproductitems.id_storepoint),

    table_all_pv as(
        select table_all_products.id_product,
               table_all_products.name_product,
               table_all_products.id_storepoint,
               table_all_products.pv_ab,
               table_all_products.pv_storepoint,
               table_all_products.pv_ab + table_all_products.pv_storepoint as pv
        from table_all_products
    ),

    table_merge_storepoint as(
        select table_all_pv.id_product,
               table_all_pv.name_product,
               table_all_pv.id_storepoint,
               coalesce(table_all_pv.pv_ab, 0) as pv_ab
        from table_all_pv
    ),

    table_result as (
        select table_merge_storepoint.id_product,
               table_merge_storepoint.name_product,
               table_merge_storepoint.id_storepoint,
               sum(table_merge_storepoint.pv_ab) as pv_ab
               from table_merge_storepoint
        group by table_merge_storepoint.id_product, table_merge_storepoint.name_product, table_merge_storepoint.id_storepoint
    )

    select table_result.id_product,
           table_result.name_product,
           sum(table_result.pv_ab) / (select count(*) from tstorepoints where isactive is true) :: double precision as pv_ab_total
    from table_result
    group by table_result.id_product, table_result.name_product
),

res as(
    select distinct all_times.id as id_product,
           all_times.name as name_product,
           all_times.id_storepoint,
           coalesce(pv.pv, 0) as pv
    from all_times
    left join pv on pv.id = all_times.id and
              pv.id_storepoint = all_times.id_storepoint
),

result as(
    select all_get_products.id_product,
           all_get_products.name_product,
           coalesce(round(coalesce(sum(res.pv) / (select count(*) from tstorepoints where isactive is true), 0), 2), 0) as pv_ab
    from all_get_products
    left join res on res.id_product = all_get_products.id_product and
              res.id_storepoint = all_get_products.id_storepoint
    group by all_get_products.id_product, all_get_products.name_product
)

select result.id_product :: bigint,
       result.name_product :: varchar,
       result.pv_ab :: double precision,
       case when result.pv_ab - table_pv_ab.pv_ab_total < 0 then 0.0
            else round(result.pv_ab - table_pv_ab.pv_ab_total, 2) :: double precision
            end as pv
from result
left join table_pv_ab on table_pv_ab.id_product = result.id_product;

end
$$;

alter function pv_function(timestamp, timestamp, integer[]) owner to sophiedata;
