package ExcelResults;

import org.testng.Assert;
import org.testng.SkipException;
import org.testng.annotations.Test;

public class TestCaseOne 
{
	 @Test
	    public void testMethodPass()
	    {
	     Assert.assertTrue(true);
	    }
	     
	    @Test
	    public void testMethodFail()
	    {
	     Assert.assertTrue(false);
	    }
	     
	    @Test
	    public void testMethodSkip()
	    {
	     throw new SkipException("Skipped Intentionally");
	    }

}
