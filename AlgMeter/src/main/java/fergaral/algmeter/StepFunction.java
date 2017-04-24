package fergaral.algmeter;

/**
 * Interface used to calculate next n value
 * 
 * @author fercarcedo
 *
 */
public interface StepFunction {
	/**
	 * Calculates the next n
	 * 
	 * @param previousN previous n
	 * @return next n given the previous value
	 */
	long nextN(long previousN);
}
