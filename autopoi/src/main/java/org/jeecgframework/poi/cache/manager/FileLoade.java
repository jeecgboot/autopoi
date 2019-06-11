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

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 文件加载类,根据路径加载指定文件
 * 
 * @author JEECG
 * @date 2014年2月10日
 * @version 1.0
 */
class FileLoade {

	private static final Logger LOGGER = LoggerFactory.getLogger(FileLoade.class);

	public byte[] getFile(String url) {
		FileInputStream fileis = null;
		ByteArrayOutputStream baos = null;
		try {
			// 先用绝对路径查询,再查询相对路径
			try {
				fileis = new FileInputStream(url);
			} catch (FileNotFoundException e) {
				String path = PoiPublicUtil.getWebRootPath(url);
				fileis = new FileInputStream(path);
			}
			baos = new ByteArrayOutputStream();
			byte[] buffer = new byte[1024];
			int len;
			while ((len = fileis.read(buffer)) > -1) {
				baos.write(buffer, 0, len);
			}
			baos.flush();
			return baos.toByteArray();
		} catch (FileNotFoundException e) {
			LOGGER.error(e.getMessage(), e);
		} catch (IOException e) {
			LOGGER.error(e.getMessage(), e);
		} finally {
			try {
				if (fileis != null)
					fileis.close();
				if (fileis != null)
					baos.close();
			} catch (IOException e) {
				LOGGER.error(e.getMessage(), e);
			}
		}
		LOGGER.error(fileis + "这个路径文件没有找到,请查询");
		return null;
	}

}
