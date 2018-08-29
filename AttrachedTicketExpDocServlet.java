package com.sdjxd.web.cq.ticket;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import com.sdjxd.web.util.ExpDoc;

public class AttrachedTicketExpDocServlet extends HttpServlet {

	private static final long serialVersionUID = 1834933L;
	private static final String templatename="attachedTemplate.ftl";
	private ExpDoc exptools = null;
	
	@Override
	public void init() throws ServletException {
		exptools = new ExpDoc();
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		doPost(request, response);
	}

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		
		String attach_data = request.getParameter("attach_data");
		if(attach_data==null){
			return;
		}
		String data = URLDecoder.decode(attach_data,"utf-8");
		Map dataMap = SetData(data,request);
		try {
			exptools.createDoc(dataMap, "每日站班会及风险控制措施检查记录表.doc",templatename);
			String outpath = trimPath(this.getClass().getResource("/")
					.toString());
			String fileName = request.getParameter("filename") == null ? "每日站班会及风险控制措施检查记录表.doc"
					: request.getParameter("filename");
			response.setHeader("Content-disposition", "attachment;filename="
					+ URLEncoder.encode(fileName, "UTF-8"));
			exptools.outPutDoc2Resp(response, outpath, fileName);
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
	}

	private Map SetData(String data, HttpServletRequest request) {
		Map<String, Object> dataMap;
		JSONObject dataobj = JSONObject.fromObject(data);
		dataMap = new HashMap<String, Object>();
		
		dataMap.put("workticketnum", dataobj.getString("num"));
		dataMap.put("workpath", dataobj.getString("workpart"));
		dataMap.put("workdate", dataobj.getString("shigongdate"));
		dataMap.put("workpeo", dataobj.getString("work_fzrname"));
		dataMap.put("safepeo", dataobj.getString("safe_jhrname"));
		dataMap.put("dangergrade", dataobj.getString("dangerChoose"));
		dataMap.put("beizhu", dataobj.getString("bz"));
		
		setSignPics(request,dataMap,dataobj.getJSONArray("allpersonimg"));
//		dataMap.put("workpeopics", dataobj.getString("allpersonimg"));
		
//		设置三交、三查数据
		setArrayData(dataMap,dataobj.getJSONArray("taskDatas"),request,"tasklist");
		setArrayData(dataMap,dataobj.getJSONArray("safemDatas"),request,"safelist");
		setArrayData(dataMap,dataobj.getJSONArray("teachDatas"),request,"techlist");
		setArrayData(dataMap,dataobj.getJSONArray("checkDatas"),request,"checklist");
		
		dataMap.put("todayctrlmeasure", dataobj.getString("checkControl"));

//		设置作业过程风险控制措施		
		setCtrlMeasures(dataMap,dataobj.getJSONArray("keyDatas"),request,"keymeasurlist");
		setCtrlMeasures(dataMap,dataobj.getJSONArray("safeSafeDatas"),request,"safemeasurelist");
		
		dataMap.put("changeissue", dataobj.getString("spotChange"));
		dataMap.put("ctrlmeasures", dataobj.getString("spotControlSafe"));
//		设置到岗到位签到列表数据
		setSignList(dataMap,dataobj.getJSONArray("signDatas"));
		
		return dataMap;

	}

	private void setSignPics(HttpServletRequest request, Map<String, Object> dataMap, JSONArray jsonArray) {
		String allsignxml = exptools.genPicsWithXml(request, jsonArray, this.getClass().getName());
		dataMap.put("workpeopics", allsignxml);
	}

	/**
	 * 设置到岗到位签到列表数据
	 * @param dataMap
	 * @param jsonArray
	 */
	private void setSignList(Map<String, Object> dataMap, JSONArray jsonArray) {
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		for(int i=0;i<jsonArray.size();i++){
			 JSONObject person = jsonArray.getJSONObject(i);
			 Map<String, String> mtemp = new HashMap<String, String>();
			 mtemp.put("unit", person.getString("parentname"));
			 mtemp.put("name", person.getString("name"));
			 mtemp.put("job", person.getString("gzname"));
			 mtemp.put("bz", person.getString("bz"));
			 list.add(mtemp);
		}
		dataMap.put("personlist", list);
	}

	/**
	 * 设置作业风险控制措施数据
	 * @param dataMap
	 * @param dataobj
	 * @param string 
	 * @param request 
	 * @param jsonArray 
	 */
	private void setCtrlMeasures(Map<String, Object> dataMap, JSONArray jsonArray, HttpServletRequest request, String tagname) {
		String img = "nochecked.png";
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		for (int i=jsonArray.size();i>0;i--) {//交任务
			Map<String, String> mtemp = new HashMap<String, String>();
			 Object datatemp = jsonArray.get(i-1);
			 JSONObject taskobj = JSONObject.fromObject(datatemp);
			 String datatext = taskobj.getString("content");
			if(taskobj.getBoolean("check")){
				img = "checked.png";
			}else{
				img = "nochecked.png";
			}
			String surepic = exptools.getCheckImg(request, img, this.getClass().getName());
			String dayneed = exptools.getCheckImg(request, "nochecked.png", this.getClass().getName());
			mtemp.put("dayneeded",dayneed);
			mtemp.put("measure",datatext);
			mtemp.put("impl", surepic);
			list.add(mtemp);
		}
		dataMap.put(tagname, list);
	}

	/**
	 * 设置三交、三查数据
	 * @param dataMap
	 * @param dataobj
	 * @param request 
	 */
	private void setArrayData(Map<String, Object> dataMap, JSONArray jarray_data, HttpServletRequest request,String tagname) {
		String img = "nochecked.png";
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		for (int i=jarray_data.size();i>0;i--) {//交任务
			Map<String, String> mtemp = new HashMap<String, String>();
			 Object datatemp = jarray_data.get(i-1);
			 JSONObject taskobj = JSONObject.fromObject(datatemp);
			 String datatext = taskobj.getString("content");
			if(taskobj.getBoolean("check")){
				img = "checked.png";
			}else{
				img = "nochecked.png";
			}
			String surepic = exptools.getCheckImg(request, img, this.getClass().getName());
			mtemp.put("data",datatext);
			mtemp.put("check", surepic);
			list.add(mtemp);
		}
		dataMap.put(tagname, list);
	}

	/**
	 * 处理文档路径
	 * 
	 * @param sourecpath
	 * @return
	 */
	private String trimPath(String sourecpath) {
		String retpath = "";
		retpath = sourecpath.replaceAll("file:/", "")
				.replaceAll("classes/", "").replaceAll("%20", " ");
		retpath = retpath.replace("webapps/cq_control/WEB-INF/", "bin/");
		return retpath;
	}

}
