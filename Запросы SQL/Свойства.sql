use RDBMZ_COPY;

;with tree(nm, id,pid, lvl,path1,path2) 
                                  as 
                                    ( 
                                    select cast( nc.COMMONTREE_NAME as varchar(1000)) COMMONTREE_NAME , nc.COMMONTREE_ID,nc.COMMONTREE_PARENT, nc.COMMONTREE_LEVEL,
                                    cast(COMMONTREE_ID as varchar(1000)) NMK_CLASSIF_ID,cast( nc.COMMONTREE_NAME as varchar(1000)) NMK_CLASSIF_NAME
                                    from COMMONTREE nc
                                    where nc.COMMONTREE_ID ='6' 
                                    union all 
                                    select cast(nc.COMMONTREE_NAME  as varchar(1000)) NMK_CLASSIF_NAME, nc.COMMONTREE_ID,nc.COMMONTREE_PARENT, nc.COMMONTREE_LEVEL ,
                                    cast(cast(path1 as varchar(1000))+'/'+cast(id as varchar(1000)) as varchar(1000))
                                                            ,   cast(cast(path2 as varchar(1000))+'/'+cast(COMMONTREE_NAME as varchar(1000)) as varchar(1000))
                                   from COMMONTREE nc 
                                    inner join tree on tree.id = nc.COMMONTREE_PARENT
                                    ) 
 
select * from
(select distinct 
par.PAR_NOTE NAME, 
PAR.PAR_CODE CODE,
	case 
	when par.PAR_TYPE='S' then 'enum'
	when par.PAR_TYPE='Y' then 'date'
	when par.PAR_TYPE='T' then 'string'
	when par.PAR_TYPE='D' then 'double'
	when par.PAR_TYPE='B' then 'bool'
	when par.PAR_TYPE='R' then 'link'
	when par.PAR_TYPE='I' then 'int'
	else par.PAR_TYPE 
	END as TYPE,
'' MEASUREENTITY,
(SELECT LEFT(DOCS.NN, LEN(DOCS.NN) - 1) AS [Expr_1] FROM  (SELECT (   SELECT   ISNULL(PAR_LIST_VAL,'')+';  '
    FROM   PAR_LIST AS DOC
      WHERE DOC.PAR_ID = l.PAR_ID order by DOC.PAR_LIST_ID
    FOR XML PATH ('')   ) AS NN  ) AS DOCS ) AS LOV, PAR.PAR_NAME DESCRIPTION, i.path2 FOLDER
                                                      from
                                                        ( select id,pid, nm,
                                   (case when substring(path1,5,1000)='' then '' else substring(path1,5,1000)+'/' end)+cast(id as varchar(1000)) path1
                                                         ,(case when substring(path2,1,1000)='' then '' else substring(path2,1,1000)+'/' end)+cast(nm as varchar(1000)) path2
                                                           from tree       group by id,pid, nm,path1,path2) i
right join PAR on par.COMMONTREE_ID=i.id
left join PAR_LIST l on l.PAR_ID=PAR.PAR_ID  ) ii
where (TYPE='S' and isnull(LOV,'') <>'') or TYPE<>'S' 
order by NAME
