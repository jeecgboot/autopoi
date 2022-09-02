/**
 * Copyright 2013-2015 JueYue (qrb.jueyue@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.cache;

import org.apache.poi.util.IOUtils;
import org.jeecgframework.poi.cache.manager.POICacheManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * 图片缓存处理
 *
 * @author liusq
 * 2022-05-27 下午4:16:32
 */
public class ImageCache {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(ImageCache.class);

    public static byte[] getImage(String imagePath) {
        InputStream                 is           = POICacheManager.getFile(imagePath);
        ByteArrayOutputStream       byteArrayOut = new ByteArrayOutputStream();
        final ByteArrayOutputStream swapStream   = new ByteArrayOutputStream();
        try {
            int ch;
            while ((ch = is.read()) != -1) {
                swapStream.write(ch);
            }
            Image         image     = Toolkit.getDefaultToolkit().createImage(swapStream.toByteArray());
            BufferedImage bufferImg = toBufferedImage(image);
            ImageIO.write(bufferImg,
                    imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length()),
                    byteArrayOut);
            return byteArrayOut.toByteArray();
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            return null;
        } finally {
            IOUtils.closeQuietly(is);
            IOUtils.closeQuietly(swapStream);
            IOUtils.closeQuietly(byteArrayOut);
        }

    }


    public static BufferedImage toBufferedImage(Image image) {
        if (image instanceof BufferedImage) {
            return (BufferedImage) image;
        }
        // This code ensures that all the pixels in the image are loaded
        image = new ImageIcon(image).getImage();
        BufferedImage bimage = null;
        GraphicsEnvironment ge = GraphicsEnvironment
                .getLocalGraphicsEnvironment();
        try {
            int                   transparency = Transparency.OPAQUE;
            GraphicsDevice        gs           = ge.getDefaultScreenDevice();
            GraphicsConfiguration gc           = gs.getDefaultConfiguration();
            bimage = gc.createCompatibleImage(image.getWidth(null),
                    image.getHeight(null), transparency);
        } catch (HeadlessException e) {
            // The system does not have a screen
        }
        if (bimage == null) {
            // Create a buffered image using the default color model
            int type = BufferedImage.TYPE_INT_RGB;
            bimage = new BufferedImage(image.getWidth(null),
                    image.getHeight(null), type);
        }
        // Copy image to buffered image
        Graphics g = bimage.createGraphics();
        // Paint the image onto the buffered image
        g.drawImage(image, 0, 0, null);
        g.dispose();
        return bimage;
    }
}
