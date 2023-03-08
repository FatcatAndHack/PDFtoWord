package com.ducway.framework.modular.zmxzApplication.controller;

import cn.stylefeng.roses.core.base.controller.BaseController;
import cn.stylefeng.roses.kernel.model.response.ResponseData;
import com.ducway.framework.modular.zmxzApplication.controller.utils.PdfToWord;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * \* User: EasonY
 * \* Date: 2023/3/6
 * \* Time: 15:41
 * \* Description:
 * \
 */
@Slf4j
@RestController
@RequestMapping("/word/system")
public class ZmxzApplicationContoller extends BaseController {

    @RequestMapping(value = "/pdfupload",method = RequestMethod.POST)
    public ResponseData pdfupload(@RequestPart("file") MultipartFile file) throws IOException {

        log.info("上传的信息：name={}，age={}，headImg={}，photos={}",
                file.getName(),file.getContentType(),file.getSize(),file.getSize());
        //1.校验
        if (file.getSize() > 100 * 1024 * 1024) {
            return ResponseData.success("文件长度过长！");
        }
        if (!file.getContentType().equalsIgnoreCase("application/pdf")) {
            return ResponseData.error("请传入PDF文件!");
        }
        //2.进行文件上传
        if (!file.isEmpty()) {
            File uploadFile = new File("D:\\cache\\");
            //文件夹不存在则创建
            if (!uploadFile.exists()) {
                uploadFile.mkdir();
            }
            uploadFile = new File("D:\\cache\\" + file.getOriginalFilename());
            //上传文件
            file.transferTo(uploadFile);
        }
        //3.拿到上传路径
        String resFileName = new PdfToWord().pdftoword("D:\\cache\\"+file.getOriginalFilename() ,file.getOriginalFilename());
        return ResponseData.success(resFileName);
    }

    @RequestMapping(value = "/downloadWord",method = RequestMethod.GET)
    public ResponseData downLoadWord(@RequestParam("name") String fileName ,
                                     HttpServletResponse response) throws IOException {

        System.out.println("测试");
        // Get your file stream from wherever.
        String fullPath = "D:/cache/";
        File downloadFile = new File(fullPath);
        FileInputStream inputStream = new FileInputStream(new java.io.File(downloadFile, fileName));
        // 设置响应头、以附件形式打开文件
        response.setHeader("content-disposition", "attachment; fileName=" + fileName);
        //获得输出流对象
        ServletOutputStream outputStream = response.getOutputStream();
        int len = 0;
        byte[] data = new byte[1024];
        while ((len = inputStream.read(data)) != -1) {
            outputStream.write(data, 0, len);
        }
        outputStream.close();
        inputStream.close();
        return ResponseData.success("保存完成");
    }

}