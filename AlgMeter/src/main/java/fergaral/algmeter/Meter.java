package fergaral.algmeter;

import java.util.Map;
import java.util.TreeMap;

/**
 * Class used to measure the execution 
 * time of a given algorithm
 * 
 * @author fercarcedo
 *
 */
public class Meter {
	private Algorithm algorithm;
	private long startN;
	private long endN;
	private StepFunction stepFunction;
	private long repetitions;

	public Meter(Algorithm algorithm, long startN, long endN, StepFunction stepFunction, long repetitions) {
		this.algorithm = algorithm;
		this.startN = startN;
		this.endN = endN;
		this.stepFunction = stepFunction;
		this.repetitions = repetitions;
	}

	/**
	 * Measures the execution time of the specified 
	 * algorithm, with n in [startN, endN], incremented 
	 * by stepN and repeated repetitions times
	 * 
	 * @return algorithm execution time as map of [n, time]
	 */
	public Map<Long, Long> run() {
		Map<Long, Long> result = new TreeMap<>();

		for (long n = startN; n <= endN; n = stepFunction.nextN(n)) {
			long totalTime = 0;
			for (long i = 0; i < repetitions; i++) {
				long startTime = System.currentTimeMillis();
				algorithm.execute(n);
				totalTime += (System.currentTimeMillis() - startTime);
			}
			long time = totalTime / repetitions;
			result.put(n, time);
		}

		return result;
	}
}
