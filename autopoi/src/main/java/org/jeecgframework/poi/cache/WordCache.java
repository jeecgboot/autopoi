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
package org.jeecgframework.poi.cache;

import java.io.InputStream;

import org.jeecgframework.poi.cache.manager.POICacheManager;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * word 缓存中心
 * 
 * @author JEECG
 * @date 2014年7月24日 下午10:54:31
 */
public class WordCache {

	private static final Logger LOGGER = LoggerFactory.getLogger(WordCache.class);

	public static MyXWPFDocument getXWPFDocumen(String url) {
		InputStream is = null;
		try {
			is = POICacheManager.getFile(url);
			MyXWPFDocument doc = new MyXWPFDocument(is);
			return doc;
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				is.close();
			} catch (Exception e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		return null;
	}

}
