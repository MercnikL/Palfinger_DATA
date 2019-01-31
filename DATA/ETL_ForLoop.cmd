for /l %%x in (1, 1, 9) do (

	IF %%x LEQ 6 CALL ETL%%x.cmd
	
	IF %%x EQU 9 CALL ETL%%x.cmd
   
)
   pause
