package decken;

public class Shape2D {
	private int x, y;
	
	Shape2D() {
		x = 0;
		y = 0;
	}
	
	Shape2D(int inX, int inY) {
		x = inX;
		y = inY;
	}
	
	int getX() {
		return x;
	}
	
	int getY() {
		return y;
	}
}
