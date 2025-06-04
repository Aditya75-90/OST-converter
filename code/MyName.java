package email.code;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.List;

public class MyName {

	public static void main(String[] args) {
		// creating object of List<String>
		List<String> arrlist = new ArrayList<String>();

		// Adding element to srclst
		arrlist.add("Ram");
		arrlist.add("Gopal");
		arrlist.add("Verma");

		// Print the list
		System.out.println("List: " + arrlist);

		// creating object of type Enumeration<String>
		Enumeration<String> e = Collections.enumeration(arrlist);

		// Print the Enumeration
		System.out.println("\nEnumeration over list: ");
		while (e.hasMoreElements()) {
			System.out.println("Value is: " + e.nextElement());
		}
	}

}
