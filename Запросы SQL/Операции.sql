use RDBMZ_COPY;

		;  with tree(nm, id,pid, lvl,path1,path2) 
                                  as 
                                    ( 
                                    select cast( nc.NMK_CLASSIF_NAME as varchar(1000)) NMK_CLASSIF_NAME , nc.NMK_CLASSIF_ID,nc.NMK_CLASSIF_PARENT, nc.NMK_CLASSIF_LEVEL,
                                    cast(NMK_CLASSIF_ID as varchar(1000)) NMK_CLASSIF_ID,cast( nc.NMK_CLASSIF_NAME as varchar(1000)) NMK_CLASSIF_NAME
                                    from NMK_CLASSIF nc
                                    where nc.NMK_CLASSIF_ID ='21' 
                                    union all 
                                    select cast(nc.NMK_CLASSIF_NAME  as varchar(1000)) NMK_CLASSIF_NAME, nc.NMK_CLASSIF_ID,nc.NMK_CLASSIF_PARENT, nc.NMK_CLASSIF_LEVEL ,
                                    cast(cast(path1 as varchar(1000))+'/'+cast(id as varchar(1000)) as varchar(1000))
									,   cast(cast(path2 as varchar(1000))+'/'+cast(NMK_CLASSIF_NAME as varchar(1000)) as varchar(1000))
                                   from NMK_CLASSIF nc 
                                    inner join tree on tree.id = nc.NMK_CLASSIF_PARENT
                                    ) 
									    select distinct n.NMK_NAME NAME,i.path2 FOLDER
								 from
								  (
								   select id,pid, nm ,
                                   (case when substring(path1,5,1000)='' then '' else substring(path1,5,1000)+'/' end)+cast(id as varchar(1000)) path1
								   ,(case when substring(path2,1,1000)='' then '' else substring(path2,1,1000)+'/' end)+cast(nm as varchar(1000)) path2
								     from tree 	 group by id,pid, nm,path1,path2
									 ) i
left join NMK n on n.NMK_CLASSIF_ID=i.id
  WHERE n.NMK_CLASSIF_ID IN (3687,9434,9436,9440,9441,3690,3692,3693,3694,9491,9513,9518,9517,9521,9504,9501,9495,9519,9503,9520,9494,9510,9507,9493,9500,9497,9511,9508,9515,9516,9502,9496,9505,9506,9498,9509,9492,9514,9499,9522,3688,3707,3702,3717,3714,3708,26971,3709,3718,3710,19944,3711,3712,3713,3715,3716,7610,23753,9645,3705,3885,12318,3888,3887,3900,3901,19676,7939,3899,5865,5972,5976,6642,18353,6643,26268,5975,5974,7804,5866,6485,6529,7801,7802,5879,3689,3699,3700,3701,3696,3695,3697,3698,3691) 
   and    n.NMK_NOTUSED='F'
order by NAME
 
