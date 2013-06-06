! PyZDDE helper macro to save JPG images of 
! a specific graphic window
CLOSEWINDOW
DATAPATH$ = $DATAPATH() # Can get the same path using zGetPath()
IMAGPATH$ = DATAPATH$ + "\IMAFiles\W02_Win02.jpg" 
EXPORTJPG 2, IMAGPATH$ , 500