! PyZDDE helper macro to save JPG images of 
! a specific graphic window
CLOSEWINDOW
DATAPATH$ = $DATAPATH() # Can get the same path using zGetPath()
IMAGPATH$ = DATAPATH$ + "\IMAFiles\W03_Win03.jpg" 
EXPORTJPG 3, IMAGPATH$ , 500