package email.code;

class teste {
	public static void main(String[] args) {
		int m = fibo(5);
		System.out.println(m);
	}

	static int fibo(int num) {
		if (num == 0) {
			return 0;
		} else if (num == 1) {
			return 1;
		} else {
			return (fibo(num - 1) + fibo(num - 2));
		}
	}
}