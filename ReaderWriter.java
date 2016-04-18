package decken;

public class ReaderWriter extends Thread {
    public static final int READER = 1;
    public static final int WRITER = 2;
    private Queue q;
    private int mode;
    public void run() {
        for (int i=0; i < 1000; i++) {
            if (mode==READER) {
                q.deQueue();
            } else if (mode==WRITER) {
                q.enQueue(new Integer(i));
            }
        }
    }
    public ReaderWriter(Queue q, int mode) {
        this.q = q;
        this.mode = mode;
    }
    public static void main(String[] args) {
    	//System.out.print("afdafdadfa");
        Queue q = new Queue(5);
        ReaderWriter r1, r2, w1, w2;
        (w1 = new ReaderWriter(q, WRITER)).start();
        (w2 = new ReaderWriter(q, WRITER)).start();
        (r1 = new ReaderWriter(q, READER)).start();
        (r2 = new ReaderWriter(q, READER)).start();
        //System.out.print("afdafdadfa");
        try {
            w1.join(); // wait until w1 complete
            w2.join(); // wait until w2 complete
            r1.join(); // wait until r1 complete
            r2.join(); // wait until r2 complete
            System.out.print("afdafdadfa");
        } catch(InterruptedException epp) {
        	System.out.print("afdafdadfa");
        }
    }
}
