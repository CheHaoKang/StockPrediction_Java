package decken;

public class Shape3D extends Shape2D{
	private int z;
	
	Shape3D() {
		super();
		z = 0;
	}
	
	Shape3D(int inX, int inY, int inZ) {
		super(inX, inY);
		z = inZ;
	}
	
	int getZ() {
		return z;
	}
}
