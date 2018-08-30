package adder;

import static org.junit.Assert.*;

import org.junit.Test;

public class addTest {

	@Test
	public void test() {
		//fail("Not yet implemented");
		adder adder = new adder();
		assertEquals(10, adder.addition(2, 8));
		
	}

}
