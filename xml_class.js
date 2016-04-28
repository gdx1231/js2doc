window.EWA = {
	VERSION : 2.6,
	/** 版本 */
	LANG : 'zhcn',
	/** 当前语言代码 */
	SHOW_ERROR : true // 是否提示执行错误
	,
	__p : window.location.pathname,
	__inc : 0
};

// 避免 http://gezz.cn/////////ex/grd.jsp 情况出现
while (EWA.__p.indexOf('//') == 0) {
	EWA.__p = EWA.__p.replace('//', '/');
	EWA.__inc++;
	if (EWA.__inc > 100) {
		break;
	}
}
/** EWA的根目录 */
EWA["CP"] = "/" + EWA.__p.split('/')[1] + "/";

var userAgent = window.navigator.userAgent.toLowerCase();
/**
 * 浏览器类型
 */
EWA.B = {
	VERSION : (userAgent.match(/.+(?:rv|it|ra|ie)[\/: ]([\d.]+)/) || [])[1],
	SAFAIR : /webkit/.test(userAgent),
	OPERA : /opera/.test(userAgent),
	IE : (/msie/.test(userAgent) || /traaaaaident/.test(userAgent)) && !/opera/.test(userAgent),
	MOZILLA : /mozilla/.test(userAgent) && !/(compatible|webkit)/.test(userAgent),
	GOOGLE : /chrome/.test(userAgent),
	PAD : /ipad|iphone|android|mobile/.test(userAgent)
};
/**
 * Xml对象
 */
function EWA_XmlClass() {
	this.XmlDoc;
	this.LoadXmlFile = function(fileUrl) {
		this._InitXmlDoc();
		this.XmlDoc.load(fileUrl);
	};
	/**
	 * 加载XML字符串
	 * 
	 * @param {Object}
	 *            strXml xml字符串
	 */
	this.LoadXml = function(strXml) {
		if (EWA.B.IE && !EWA.B.MOZILLA) { //IE<11
			this._InitXmlDoc();
			this.XmlDoc.loadXML(strXml);
		} else {
			this.XmlDoc = null;
			this.XmlDoc = new DOMParser().parseFromString(strXml, "text/xml");
		}
	};
	/**
	 * 获取XML节点的XML
	 * 
	 * @param {Object}
	 *            node 节点
	 * @return 如果Node不为空，返回Node的XML，否则返回DOC的XML
	 */
	this.GetXml = function(node) {
		if (node == null) {
			node = this.XmlDoc;
		}
		if (!EWA.B.IE) {
			return (new XMLSerializer).serializeToString(node);
		} else {
			return node.xml;
		}
	};
	/**
	 * 获取属性值
	 * 
	 * @param {Object}
	 *            path 路径
	 * @param {Object}
	 *            attributeName 属性名称
	 * @param {Object}
	 *            element 当前对象
	 * @return (String) 属性值
	 */
	this.GetAttributeValue = function(path, attributeName, element) {
		var a = this.GetElement(path, element);
		return a.getAttribute(attributeName);
	};
	/**
	 * 获取对象的Text值
	 * 
	 * @param {Object}
	 *            element 对象
	 * @return (String) 值
	 */
	this.GetElementText = function(element) {
		if (element == null)
			return null;
		if (element.childNodes.length > 0) {
			var s1 = "";
			for (var i = 0; i < element.childNodes.length; i++) {
				s1 += element.childNodes[i].nodeValue;
			}
			return s1;
		} else {
			return null;
		}
	};
	/**
	 * 获取对象属性值
	 * 
	 * @param {Object}
	 *            element 对象
	 * @param {Object}
	 *            attributeName 属性名称
	 */
	this.GetElementAttribute = function(element, attributeName) {
		return element.getAttribute(attributeName);
	};
	/**
	 * 获取对象的Text值
	 * 
	 * @param {Object}
	 *            path 路径
	 * @param {Object}
	 *            element 对象
	 */
	this.GetText = function(path, element) {
		var a = this.GetElement(path, element);
		return this.GetElementText(a);
	};
	/**
	 * 获取对象列表
	 * 
	 * @param {Object}
	 *            path 路径
	 * @param {Object}
	 *            element 对象
	 * @return 对象列表数组
	 */
	this.GetElements = function(path, element) {
		var paths = path.split("/");
		var a;
		if (element == null) {
			a = this.XmlDoc;
		} else {
			a = element;
		}
		for (var i = 0; i < paths.length - 1; i += 1) {// for firefox
			var b = a.getElementsByTagName(paths[i]);
			if (b == null || b.length == 0) {
				return null;
			}
			a = b[0];
			b = null;
		}
		if (a == null || a.length == 0) {
			return null;
		} else {
			return a.getElementsByTagName(paths[paths.length - 1]);
		}
	};
	/**
	 * 获取对象
	 * 
	 * @param {String}
	 *            path 路径
	 * @param {Object}
	 *            element 对象
	 */
	this.GetElement = function(path, element) {
		var a = this.GetElements(path, element);
		if (a == null) {
			return null;
		} else {
			return a[0];
		}
	};
	/**
	 * 对象属性赋值
	 * 
	 * @param {String}
	 *            attName 属性名称
	 * @param {String}
	 *            attValue 属性值
	 * @param {Object}
	 *            element 对象
	 */
	this.SetAttribute = function(attName, attValue, element) {
		if (attValue == null) {
			element.setAttribute(attName, "");
		} else {
			element.setAttribute(attName, attValue);
		}
	};
	/**
	 * 生成或获取节点
	 * 
	 * @param {String}
	 *            path 路径
	 * @param {Object}
	 *            element 当前节点
	 */
	this.GetOrCreateElement = function(path, element) {
		var paths = path.split("/");
		var a;
		if (element == null) {
			a = this.XmlDoc;
		} else {
			a = element;
		}
		for (var i = 0; i < paths.length; i += 1) {// for firefox
			var b = a.getElementsByTagName(paths[i]);
			if (b == null || b.length == 0) {
				a = this.NewChild(paths[i], a);
				b = null;
			} else {
				a = b[0];
				b = null;
			}
		}
		return a;
	}
	/**
	 * 新字节点
	 * 
	 * @param {String}
	 *            tagName 节点Tag名称
	 * @param {Object}
	 *            elementParent 父节点
	 */
	this.NewChild = function(tagName, elementParent) {
		var ele = this.XmlDoc.createElement(tagName);
		elementParent.appendChild(ele);
		return ele;
	};
	this.NewChilds = function(tagNames, elementParent) {
		var tags = tagNames.split('/');
		var eleP = elementParent;
		for (var i = 0; i < tags.length; i++) {
			var eleP = this.NewChild(tags[i].trim(), eleP)
		}
		return eleP;
	};
	/**
	 * 设置节点值
	 * 
	 * @param {Object}
	 *            text 值
	 * @param {Object}
	 *            element 节点
	 */
	this.SetText = function(text, element) {
		element.text = text;
	};
	/**
	 * 新增文本节点
	 * 
	 * @param {Object}
	 *            text
	 * @param {Object}
	 *            element
	 */
	this.AppendText = function(text, element) {
		var node = this.XmlDoc.createTextNode(text);
		element.appendChild(node);
	};
	/**
	 * 赋值CDATA节点
	 * 
	 * @param {String}
	 *            text
	 * @param {Object}
	 *            element
	 */
	this.SetCData = function(text, element) {
		if (text == null) {
			return;
		}
		var c = this.XmlDoc.createCDATASection(text);
		element.appendChild(c);
	};
	/**
	 * 初始化XML Document
	 */
	this._InitXmlDoc = function() {
		if (window.ActiveXObject) {
			var xmls = [ "MSXML4.DOMDocument", "MSXML3.DOMDocument", "MSXML2.DOMDocument", "MSXML.DOMDocument",
					"Microsoft.XmlDom" ];
			for (var i = 0; i < xmls.length; i++) {
				try {
					this.XmlDoc = new ActiveXObject(xmls[i]);
				} catch (e) {
					this.XmlDoc = null;
				}
				if (this.XmlDoc != null) {
					break;
				}
			}
		} else {
			// 创建FIREFOX下XML文档对象
			this.XmlDoc = document.implementation.createDocument("", "doc", null);
		}
		try {
			this.XmlDoc.async = false;
		} catch (e) {

		}
	};
	this._InitXmlDoc();
}
