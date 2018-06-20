    '3D MATRIX info
	'First 9 lines of the matrix defines a 3D rotating base
	'reference is from Parent item. Pbase(P,x,y,z)
	
	
	Set move1 = product2.Move
    Set move1 = move1.MovableObject

	'Coordinates of vector u in Pbase
    movematrix(0) = 1
    movematrix(1) = 0
    movematrix(2) = 0

	'Coordinates of vector v in Pbase
    movematrix(3) = 0
    movematrix(4) = A
    movematrix(5) = A
	'Coordinates of vector w in Pbase
    movematrix(6) = 0
    movematrix(7) = -A
    movematrix(8) = A
		
	'Transaltion matrix over (P,x,y,z)
    movematrix(9) = 0#
    movematrix(10) = 0#
    movematrix(11) = 0#
    
	'A new base for the child Cbase(C,u,v,w) is now defined.
	
	move1.Apply movematrix
	
	'Extra rules may apply
	
