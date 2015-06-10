! PyZDDE helper macro to save JPG images of 
! a specific graphic window
CLOSEWINDOW
DATAPATH$ = $DATAPATH() # Can get the same path using zGetPath()
IMAGPATH$ = DATAPATH$ + "\IMAFiles\W04_Win05.jpg" 
EXPORTJPG 5, IMAGPATH$ , 500