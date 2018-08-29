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
	 * ����doc�ĵ�
	 * @param dataMap
	 * @param fileName  such as a.doc
	 * @param templateName   such as  b.ftl
	 * @throws UnsupportedEncodingException
	 */
	public void createDoc(Map<String, Object> dataMap, String fileName,String templateName)
			throws UnsupportedEncodingException {
		// dataMap Ҫ����ģ���������ļ�
		// ����ģ��װ�÷�����·��,FreeMarker֧�ֶ���ģ��װ�ط�����������servlet��classpath�����ݿ�װ�أ�
		// �������ǵ�ģ���Ƿ���template������
		configuration.setClassForTemplateLoading(this.getClass(), TEMPLATE_PATH);
		Template t = null;
		try {
			// test.ftlΪҪװ�ص�ģ��
			t = configuration.getTemplate(templateName);
		} catch (IOException e) {
			e.printStackTrace();
		}
		// ����ĵ�·��������
		File outFile = new File(fileName);
		Writer out = null;
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(outFile);
			OutputStreamWriter oWriter = new OutputStreamWriter(fos, "UTF-8");
			// ����ط������ı��벻�ɻ�ȱ��ʹ��main������������ʱ��Ӧ�ÿ��ԣ����������web���󵼳�ʱ������word�ĵ��ͻ�򲻿������Ұ�XML�ļ�������Ҫ�Ǳ����ʽ����ȷ���޷�������
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
	 *����base64�����url����ģ���е�ͼƬ
	 * @param base64url
	 * @param width  �԰�Ϊ��λ
	 * @param height �԰�Ϊ��λ
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
	 * ��ȡ��ҵ�ر������Ƿ�ѡ�е�ͼƬ
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
	 * ��ȡbase64�����url
	 * @param url
	 * @param request
	 * @param classpkgtname
	 * @return
	 */
	private String getBase64Url(URL url, HttpServletRequest request,String classpkgtname) {
		String base64 = "";
		// String extName = getExtName(url.toString());
		// ������
		HttpURLConnection conn = null;
		try {
			conn = (HttpURLConnection) url.openConnection();
			// ��������ʽΪ"GET"
			conn.setRequestMethod("GET");
			// ��ʱ��Ӧʱ��Ϊ5��
			conn.setConnectTimeout(5 * 1000);
			// ͨ����������ȡͼƬ����
			if (conn.getResponseCode() == 404) {
				String str404 = getNotFindPic(request,classpkgtname);
				// extName = getExtName(str404);
				URL url404 = new URL(str404);
				conn = (HttpURLConnection) url404.openConnection();
			}
			InputStream inStream = conn.getInputStream();
			// �õ�ͼƬ�Ķ��������ݣ��Զ����Ʒ�װ�õ����ݣ�����ͨ����
			ByteArrayOutputStream outStream = new ByteArrayOutputStream();
			// ����һ��Buffer�ַ���
			byte[] buffer = new byte[1024];
			// ÿ�ζ�ȡ���ַ������ȣ����Ϊ-1������ȫ����ȡ���
			int len = 0;
			// ʹ��һ����������buffer������ݶ�ȡ����
			while ((len = inStream.read(buffer)) != -1) {
				// ���������buffer��д�����ݣ��м����������ĸ�λ�ÿ�ʼ����len�����ȡ�ĳ���
				outStream.write(buffer, 0, len);
			}
			
			byte[] data = outStream.toByteArray();
			// ���ֽ�����Base64����
			// �ر���
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

		return base64;// ����Base64��������ֽ������ַ���
	}
	
	
	/**
	 * δ�ҵ�ͼƬʱ����δ�ҵ���ͼƬ��ʾ����ֹ�հ�
	 * @param request ��ǰ������
	 * @param classpkgtname ��ǰ��servletʵ���İ��� this.getClass().getName();
	 * @return "ͼƬ��ʧ"ʾ��ͼƬ��·�� 
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
	 * ��ȡ xml��ʽ��doc�ĵ��е�ͼƬԪ��
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
	 * û��ͼƬ��ʱ����ʾ��Ĭ��ͼƬ
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
