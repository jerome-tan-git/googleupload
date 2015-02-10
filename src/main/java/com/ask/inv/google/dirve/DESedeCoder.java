package com.ask.inv.google.dirve;
import java.security.SecureRandom;  

import javax.crypto.Cipher;  
import javax.crypto.SecretKey;  
import javax.crypto.SecretKeyFactory;  
import javax.crypto.spec.DESKeySpec;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.security.Key;
import javax.crypto.Cipher;
import javax.crypto.KeyGenerator;
import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.DESedeKeySpec;
import com.sun.org.apache.xml.internal.security.utils.Base64;
public class DESedeCoder  {  
	  
    /**
     * 密钥算法
     * */
    public static final String KEY_ALGORITHM = "DESede";

    /**
     * 加密/解密算法/工作模式/填充方式
     * */
    public static final String CIPHER_ALGORITHM = "DESede/ECB/PKCS5Padding";

    /**
     * 
     * 生成密钥
     * 
     * @return byte[] 二进制密钥
     * */
    public static byte[] initkey() throws Exception {

        // 实例化密钥生成器
        KeyGenerator kg = KeyGenerator.getInstance(KEY_ALGORITHM);
        // 初始化密钥生成器
        kg.init(168);
        // 生成密钥
        SecretKey secretKey = kg.generateKey();
        // 获取二进制密钥编码形式
        
        byte[] key = secretKey.getEncoded();
        BufferedOutputStream keystream = 
                new BufferedOutputStream(new FileOutputStream("DESedeKey.dat"));
        keystream.write(key, 0, key.length);
        keystream.flush();
        keystream.close();
        
        return key;
    }

    /**
     * 转换密钥
     * 
     * @param key
     *            二进制密钥
     * @return Key 密钥
     * */
    public static Key toKey(byte[] key) throws Exception {
        // 实例化Des密钥
        DESedeKeySpec dks = new DESedeKeySpec(key);
        // 实例化密钥工厂
        SecretKeyFactory keyFactory = SecretKeyFactory
                .getInstance(KEY_ALGORITHM);
        // 生成密钥
        SecretKey secretKey = keyFactory.generateSecret(dks);
        return secretKey;
    }

    /**
     * 加密数据
     * 
     * @param data
     *            待加密数据
     * @param key
     *            密钥
     * @return byte[] 加密后的数据
     * */
    public static byte[] encrypt(byte[] data, byte[] key) throws Exception {
        // 还原密钥
        Key k = toKey(key);
        // 实例化
        Cipher cipher = Cipher.getInstance(CIPHER_ALGORITHM);
        // 初始化，设置为加密模式
        cipher.init(Cipher.ENCRYPT_MODE, k);
        // 执行操作
        return cipher.doFinal(data);
    }

    /**
     * 解密数据
     * 
     * @param data
     *            待解密数据
     * @param key
     *            密钥
     * @return byte[] 解密后的数据
     * */
    public static byte[] decrypt(byte[] data, byte[] key) throws Exception {
        // 欢迎密钥
        Key k = toKey(key);
        // 实例化
        Cipher cipher = Cipher.getInstance(CIPHER_ALGORITHM);
        // 初始化，设置为解密模式
        cipher.init(Cipher.DECRYPT_MODE, k);
        // 执行操作
        return cipher.doFinal(data);
    }

    /**
     * 进行加解密的测试
     * 
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String str = "123!@#QWEqweE";
        System.out.println("原文：" + str);
        // 初始化密钥
        byte[] key = Base64.decode("/imPE5FUDowZE4ltFWHsSVekH7nav0zl");
        System.out.println("密钥：" + Base64.encode(key));
        // 加密数据
        byte[] data = DESedeCoder.encrypt(str.getBytes(), key);
        String mi = Base64.encode(data);
        System.out.println("加密后：" + mi);
        // 解密数据
//        String mi="O2mUilnGw84GwNk9DpzVOQ==";
//        data = DESedeCoder.decrypt(Base64.decode(mi), Base64.decode("/imPE5FUDowZE4ltFWHsSVekH7nav0zl"));
        System.out.println("解密后：" + new String(data));
    }
}  