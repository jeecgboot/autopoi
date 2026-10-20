package org.jeecgframework.poi.util;

import org.junit.Test;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicInteger;

import static org.junit.Assert.assertEquals;

/**
 * PoiFunctionUtil 的 formatDate / formatNumber 在多线程下必须得到正确结果。
 * 内部缓存了共享的 SimpleDateFormat / DecimalFormat，而这两个类都是线程不安全的。
 *
 * @since 2026-10-20 issues/52 SimpleDateFormat有线程安全问题
 */
public class PoiFunctionUtilThreadSafeTest {

	private static final int THREADS = 8;
	private static final int LOOPS = 3000;

	/** 并发执行 task，返回出错（结果不一致或抛异常）的总次数 */
	private static int runConcurrently(final Task task) throws Exception {
		ExecutorService pool = Executors.newFixedThreadPool(THREADS);
		final AtomicInteger wrong = new AtomicInteger();
		List<Future<?>> futures = new ArrayList<Future<?>>();
		for (int t = 0; t < THREADS; t++) {
			final int seed = t;
			futures.add(pool.submit(new Callable<Void>() {
				@Override
				public Void call() {
					for (int i = 0; i < LOOPS; i++) {
						try {
							if (!task.run(seed, i)) {
								wrong.incrementAndGet();
							}
						} catch (Exception e) {
							wrong.incrementAndGet();
						}
					}
					return null;
				}
			}));
		}
		for (Future<?> f : futures) {
			f.get();
		}
		pool.shutdown();
		return wrong.get();
	}

	private interface Task {
		/** @return 结果是否正确 */
		boolean run(int seed, int i) throws Exception;
	}

	private static Date dateOf(int seed, int i) {
		Calendar c = Calendar.getInstance();
		c.set(2000 + seed, i % 12, 1 + i % 28, i % 24, i % 60, i % 60);
		return c.getTime();
	}

	private static void assertDateFormatSafe(final String pattern) throws Exception {
		final ThreadLocal<SimpleDateFormat> expected = new ThreadLocal<SimpleDateFormat>() {
			@Override
			protected SimpleDateFormat initialValue() {
				return new SimpleDateFormat(pattern);
			}
		};
		int wrong = runConcurrently(new Task() {
			@Override
			public boolean run(int seed, int i) {
				Date d = dateOf(seed, i);
				return expected.get().format(d).equals(PoiFunctionUtil.formatDate(d, pattern));
			}
		});
		assertEquals(pattern + " 并发格式化出错次数", 0, wrong);
	}

	@Test
	public void formatDateWithTimePatternShouldBeThreadSafe() throws Exception {
		assertDateFormatSafe("yyyy-MM-dd HH:mm:ss");
	}

	@Test
	public void formatDateWithDayPatternShouldBeThreadSafe() throws Exception {
		assertDateFormatSafe("yyyy-MM-dd");
	}

	@Test
	public void formatDateWithoutSecondPatternShouldBeThreadSafe() throws Exception {
		assertDateFormatSafe("yyyy-MM-dd HH:mm");
	}

	@Test
	public void formatDateWithCustomPatternShouldBeThreadSafe() throws Exception {
		assertDateFormatSafe("yyyy/MM/dd");
	}

	@Test
	public void formatNumberShouldBeThreadSafe() throws Exception {
		final String[] patterns = {"###.00", "###.000", "0.0"};
		for (final String pattern : patterns) {
			final ThreadLocal<DecimalFormat> expected = new ThreadLocal<DecimalFormat>() {
				@Override
				protected DecimalFormat initialValue() {
					return new DecimalFormat(pattern);
				}
			};
			int wrong = runConcurrently(new Task() {
				@Override
				public boolean run(int seed, int i) {
					double value = seed * 1000 + i + 0.123456;
					return expected.get().format(value).equals(PoiFunctionUtil.formatNumber(value, pattern));
				}
			});
			assertEquals(pattern + " 并发格式化出错次数", 0, wrong);
		}
	}

	@Test
	public void formatResultShouldBeCorrectInSingleThread() {
		Calendar c = Calendar.getInstance();
		c.set(2026, Calendar.OCTOBER, 20, 8, 5, 3);
		Date d = c.getTime();
		assertEquals("2026-10-20", PoiFunctionUtil.formatDate(d, "yyyy-MM-dd"));
		assertEquals("2026-10-20 08:05:03", PoiFunctionUtil.formatDate(d, "yyyy-MM-dd HH:mm:ss"));
		assertEquals("2026-10-20 08:05", PoiFunctionUtil.formatDate(d, "yyyy-MM-dd HH:mm"));
		assertEquals("12.35", PoiFunctionUtil.formatNumber(12.3456, "###.00"));
		assertEquals("12.346", PoiFunctionUtil.formatNumber(12.3456, "###.000"));
		assertEquals("", PoiFunctionUtil.formatDate(null, "yyyy-MM-dd"));
		assertEquals("", PoiFunctionUtil.formatNumber(null, "###.00"));
	}
}
