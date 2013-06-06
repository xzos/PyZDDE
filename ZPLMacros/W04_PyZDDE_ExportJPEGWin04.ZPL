! PyZDDE helper macro to save JPG images of 
! a specific graphic window
CLOSEWINDOW
DATAPATH$ = $DATAPATH() # Can get the same path using zGetPath()
IMAGPATH$ = DATAPATH$ + "\IMAFiles\W04_Win04.jpg" 
EXPORTJPG 4, IMAGPATH$ , 500