package com.sdjxd.web.util;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.poi.hssf.record.formula.functions.T;

import sun.misc.BASE64Encoder;

import com.sdjxd.hussar.core.utils.Guid;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class ExpDoc {
	private static final String PIC_WIDTH = "80.75";
	private static final String PIC_HEIGHT = "108.75";
	private static final String TEMPLATE_PATH = "/template";
	private Configuration configuration = null;
	public ExpDoc() {
		configuration = new Configuration();
		configuration.setDefaultEncoding("utf-8");
	}
	
	/**
	 * @param response 
	 * @param outpath  tocat/bin
	 * @param fileName  a.doc
	 */
	public void outPutDoc2Resp(HttpServletResponse response,String outpath,String fileName){
		FileInputStream fis=null;
		OutputStream oa=null;
		try {
			fis = new FileInputStream(outpath + fileName);
			oa = response.getOutputStream();
			byte[] b = new byte[1024];
			int i = 0;
			while ((i = fis.read(b)) > 0) {
				oa.write(b, 0, i);
			}
			fis.close();
			oa.flush();
			oa.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 创建doc文档
	 * @param dataMap
	 * @param fileName  such as a.doc
	 * @param templateName   such as  b.ftl
	 * @throws UnsupportedEncodingException
	 */
	public void createDoc(Map<String, Object> dataMap, String fileName,String templateName)
			throws UnsupportedEncodingException {
		// dataMap 要填入模本的数据文件
		// 设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，
		// 这里我们的模板是放在template包下面
		configuration.setClassForTemplateLoading(this.getClass(), TEMPLATE_PATH);
		Template t = null;
		try {
			// test.ftl为要装载的模板
			t = configuration.getTemplate(templateName);
		} catch (IOException e) {
			e.printStackTrace();
		}
		// 输出文档路径及名称
		File outFile = new File(fileName);
		Writer out = null;
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(outFile);
			OutputStreamWriter oWriter = new OutputStreamWriter(fos, "UTF-8");
			// 这个地方对流的编码不可或缺，使用main（）单独调用时，应该可以，但是如果是web请求导出时导出后word文档就会打不开，并且包XML文件错误。主要是编码格式不正确，无法解析。
			// out = new BufferedWriter(new OutputStreamWriter(new
			// FileOutputStream(outFile)));
			out = new BufferedWriter(oWriter);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}

		try {
			t.process(dataMap, out);
			out.close();
			fos.close();
		} catch (TemplateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 
	 *根据base64编码的url生成模板中的图片
	 * @param base64url
	 * @param width  以磅为单位
	 * @param height 以磅为单位
	 * @return
	 */
	private StringBuffer getPicItemWithBase64(String base64url,String imgname,String width,String height) {
		StringBuffer pictemp = new StringBuffer();
		width=width==null?this.PIC_WIDTH:width;
		width=height==null?this.PIC_HEIGHT:height;
		String shaptypeid = Guid.create();
		String shapid = Guid.create();
		pictemp.append("<w:pict>");
		pictemp.append("<v:shapetype id=\""+shaptypeid+"\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">");
		pictemp.append("<v:stroke joinstyle=\"miter\"/>");
		pictemp.append("<v:formulas> ");
		pictemp.append("<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>");
		pictemp.append("<v:f eqn=\"sum @0 1 0\"/>");
		pictemp.append("<v:f eqn=\"sum 0 0 @1\"/>");
		pictemp.append("<v:f eqn=\"prod @2 1 2\"/>");
		pictemp.append("<v:f eqn=\"prod @3 21600 pixelWidth\"/>");
		pictemp.append("<v:f eqn=\"prod @3 21600 pixelHeight\"/>");
		pictemp.append("<v:f eqn=\"sum @0 0 1\"/>");
		pictemp.append("<v:f eqn=\"prod @6 1 2\"/>");
		pictemp.append("<v:f eqn=\"prod @7 21600 pixelWidth\"/>");
		pictemp.append("<v:f eqn=\"sum @8 21600 0\"/>");
		pictemp.append("<v:f eqn=\"prod @7 21600 pixelHeight\"/>");
		pictemp.append("<v:f eqn=\"sum @10 21600 0\"/>");
		pictemp.append("</v:formulas>");
		pictemp.append("<v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
		pictemp.append("<o:lock v:ext=\"edit\" aspectratio=\"t\"/>");
		pictemp.append("</v:shapetype>");
		pictemp.append("<w:binData w:name=\"wordml://"+imgname+"\">");
		pictemp.append(base64url);
		pictemp.append("</w:binData> ");
		pictemp.append("<v:shape id=\""+shapid+"\" type=\"#_x0000_t75\" style=\"width:"+width+"pt;height:"+height+"pt\">");
		pictemp.append("<v:imagedata src=\"wordml://"+imgname+"\" o:title=\"nopics\"/>");
		pictemp.append("</v:shape>");
		pictemp.append("</w:pict>");
		return pictemp;
	}
	/**
	 * 获取作业必备条件是否选中的图片
	 * @param request
	 * @param imgname
	 * @param classpkgtname (com.sd.jxd.xxxServlet) </br>
	 * 	the value of the method  this.getClass().getName()  returned 
	 * @return
	 */
	
	public String getCheckImg(HttpServletRequest request, String imgname,String classpkgtname) {
		String base64url = "";
		StringBuffer checkXml = new StringBuffer();
		StringBuffer requrl = request.getRequestURL();
		String servlatname = classpkgtname.substring(classpkgtname.lastIndexOf(".")+1);
		String procontext = requrl.substring(0, requrl.indexOf(servlatname));
		String imgurl = procontext + "cq/img/"+imgname;
		try {
			URL url =  new URL(imgurl);
			String picname = getPicNameFromurl(imgurl);
			base64url = getBase64Url(url, request,classpkgtname);
			checkXml = getPicItemWithBase64(base64url,picname,"9","9");
		} catch (MalformedURLException e) {
			e.printStackTrace();
		}
		return checkXml.toString();
	}
	
	private String getPicNameFromurl(String path) {
		String filename = path.substring(path.lastIndexOf("/")+1);
		return filename;
	}
	
	/**
	 * 获取base64编码的url
	 * @param url
	 * @param request
	 * @param classpkgtname
	 * @return
	 */
	private String getBase64Url(URL url, HttpServletRequest request,String classpkgtname) {
		String base64 = "";
		// String extName = getExtName(url.toString());
		// 打开链接
		HttpURLConnection conn = null;
		try {
			conn = (HttpURLConnection) url.openConnection();
			// 设置请求方式为"GET"
			conn.setRequestMethod("GET");
			// 超时响应时间为5秒
			conn.setConnectTimeout(5 * 1000);
			// 通过输入流获取图片数据
			if (conn.getResponseCode() == 404) {
				String str404 = getNotFindPic(request,classpkgtname);
				// extName = getExtName(str404);
				URL url404 = new URL(str404);
				conn = (HttpURLConnection) url404.openConnection();
			}
			InputStream inStream = conn.getInputStream();
			// 得到图片的二进制数据，以二进制封装得到数据，具有通用性
			ByteArrayOutputStream outStream = new ByteArrayOutputStream();
			// 创建一个Buffer字符串
			byte[] buffer = new byte[1024];
			// 每次读取的字符串长度，如果为-1，代表全部读取完毕
			int len = 0;
			// 使用一个输入流从buffer里把数据读取出来
			while ((len = inStream.read(buffer)) != -1) {
				// 用输出流往buffer里写入数据，中间参数代表从哪个位置开始读，len代表读取的长度
				outStream.write(buffer, 0, len);
			}
			
			byte[] data = outStream.toByteArray();
			// 对字节数组Base64编码
			// 关闭流
			inStream.close();
			outStream.close();
			conn.disconnect();
			BASE64Encoder encoder = new BASE64Encoder();
			base64 = encoder.encode(data);
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			conn.disconnect();
		}

		return base64;// 返回Base64编码过的字节数组字符串
	}
	
	
	/**
	 * 未找到图片时设置未找到的图片显示，防止空白
	 * @param request 当前的请求
	 * @param classpkgtname 当前的servlet实例的包名 this.getClass().getName();
	 * @return "图片丢失"示例图片的路径 
	 */
	private String getNotFindPic(HttpServletRequest request,String classpkgtname) {
		String retNotFindstr = "";
		StringBuffer nopicsxml = new StringBuffer();
		StringBuffer requrl = request.getRequestURL();
		String servlatname = classpkgtname.substring(classpkgtname.lastIndexOf(".")+1);
		String procontext = requrl.substring(0, requrl.indexOf(servlatname));
		retNotFindstr = procontext + "cq/img/losedpics.png";
		return retNotFindstr;
	}

	/**
	 * 获取 xml格式的doc文档中的图片元素
	 * @param request
	 * @param jsonArray
	 * @param classpkgname  (com.sd.jxd.xxxServlet) </br>
	 * 		this.getClass().getName()
	 * @return
	 */
	public String genPicsWithXml(HttpServletRequest request, JSONArray jsonArray,String classpkgname) {
		StringBuffer retXml = new StringBuffer();
		StringBuffer pictemp = new StringBuffer();
		if (jsonArray.size() == 0) {
			retXml = genNopicsXml(request,classpkgname);
		}
		for (int i = 0; i < jsonArray.size(); i++) {
			Object temp = jsonArray.get(i);
			JSONObject pathtemp = JSONObject.fromObject(temp);
			String path = pathtemp.getString("filepath");
			String  picname = getPicNameFromurl(path);
			URL url = null;
			try {
				url = new URL(path);
			} catch (MalformedURLException e) {
				e.printStackTrace();
			}
			String base64url = getBase64Url(url, request, classpkgname);
			pictemp = getPicItemWithBase64(base64url,picname,"","");
			retXml.append(pictemp);
		}

		return retXml.toString();
	}
	/**
	 * 没有图片的时候显示的默认图片
	 * @param request
	 * @param classpkgname {@link genPicsWithXml}
	 * @return
	 */
	private StringBuffer genNopicsXml(HttpServletRequest request,String classpkgname) {
		StringBuffer nopicsxml = new StringBuffer();
		StringBuffer requrl = request.getRequestURL();
		String servletname = classpkgname.substring(classpkgname.lastIndexOf(".")+1);
		String procontext = requrl.substring(0, requrl.indexOf(servletname));
		String nopicurl = procontext + "cq/img/nopics.png";
		String picname = getPicNameFromurl(nopicurl);
		String nopicsbase64 = "";
		try {
			URL url = new URL(nopicurl);
			nopicsbase64 = getBase64Url(url, request,classpkgname);
		} catch (MalformedURLException e) {
			e.printStackTrace();
		}
		nopicsxml = getPicItemWithBase64(nopicsbase64,picname,"","");
		return nopicsxml;
	}
}
