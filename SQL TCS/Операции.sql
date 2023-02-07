use RDBMZ_COPY;

SELECT distinct NMK.NMK_NAME,NMK_CLASSIF.NMK_CLASSIF_NAME
  FROM NMK NMK
     LEFT JOIN NMK_CLASSIF NMK_CLASSIF ON  (NMK.NMK_CLASSIF_ID = NMK_CLASSIF.NMK_CLASSIF_ID)  
         LEFT JOIN MESURIMENT ON  (NMK.NMK_BASE_MESUR = MESURIMENT.MESUR_ID) 
          WHERE NMK.NMK_CLASSIF_ID IN (3687,9434,9436,9440,9441,3690,3692,3693,3694,9491,9513,9518,9517,9521,9504,9501,9495,9519,9503,9520,9494,9510,9507,9493,9500,9497,9511,9508,9515,9516,9502,9496,9505,9506,9498,9509,9492,9514,9499,9522,3688,3707,3702,3717,3714,3708,26971,3709,3718,3710,19944,3711,3712,3713,3715,3716,7610,23753,9645,3705,3885,12318,3888,3887,3900,3901,19676,7939,3899,5865,5972,5976,6642,18353,6643,26268,5975,5974,7804,5866,6485,6529,7801,7802,5879,3689,3699,3700,3701,3696,3695,3697,3698,3691) 
          and NMK.NMK_NOTUSED='F'
