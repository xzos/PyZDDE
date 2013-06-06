! PyZDDE helper macro to save JPG images of 
! a specific graphic window
CLOSEWINDOW
DATAPATH$ = $DATAPATH() # Can get the same path using zGetPath()
IMAGPATH$ = DATAPATH$ + "\IMAFiles\W01_Win01.jpg" 
EXPORTJPG 1, IMAGPATH$ , 500