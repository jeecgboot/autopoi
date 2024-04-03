/**
 * Copyright 2013-2015 JEECG (jeecgos@163.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.cache.manager;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

/**
 * 缓存管理
 * 
 * @author JEECG
 * @date 2014年2月10日
 * @version 1.0
 */
public final class POICacheManager {

	private static final Logger LOGGER = LoggerFactory.getLogger(POICacheManager.class);

	private static LoadingCache<String, byte[]> loadingCache;

	static {
		loadingCache = CacheBuilder.newBuilder().expireAfterWrite(7, TimeUnit.DAYS).maximumSize(50).build(new CacheLoader<String, byte[]>() {
			@Override
			public byte[] load(String url) throws Exception {
				return new FileLoade().getFile(url);
			}
		});
	}

	public static InputStream getFile(String id) {
		try {
			// 复杂数据,防止操作原数据
			byte[] result = Arrays.copyOf(loadingCache.get(id), loadingCache.get(id).length);
			return new ByteArrayInputStream(result);
		} catch (ExecutionException e) {
			LOGGER.error(e.getMessage(), e);
		}
		return null;
	}

	//update-begin---author:chenrui ---date:20240403  for：[issue/#5933]增加清除缓存方法------------
    /**
     * 清除所有缓存
     *
     * @author chenrui
     * @date 2024/4/3 11:46
     */
    public static void cleanAll() {
        loadingCache.invalidateAll();
    }

    /**
     * 清除缓存
     *
     * @param id 缓存key
     * @author chenrui
     * @date 2024/4/3 11:46
     */
    public static void clean(String id) {
        loadingCache.invalidate(id);
    }
	//update-end---author:chenrui ---date:20240403  for：[issue/#5933]增加清除缓存方法------------

}
