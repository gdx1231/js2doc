function EWA_OdtDocWordClass() {
	this.docks = [];
	this.pics = [];
	this.liNum = 0;// numering.xml defined
	this.start = false;
	this.walker = function(obj) {

		this.fixTable();
		this.walker1(obj);
		this.unFixTable();
	};
	this.walker1 = function(obj) {

		if (obj.nodeType == 3) {
			this.walkerTxt(obj)
		} else if (obj.nodeType == 1) {
			this.walkerEle(obj)
		}
	};
	this.walkerEle = function(obj) {
		if (obj.style.display == 'none') {
			return;
		}
		var t = obj.tagName;
		var endPop = false;
		var o;
		if (t == 'DIV'
				&& (obj.className.indexOf('page-next-before') >= 0 || obj.className
						.indexOf('page-next-after') >= 0)) {
			// 分页符号
			o = this.createBreak();
			var p = this.docks[0];
			console.log('break',o)
			p.appendChild(o);
			endPop = true;
		} else if (t == 'UL' || t == 'OL') {
			o = this.createUl(obj);
			var p = this.getDockTable();
			p.appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'LI') {
			o = this.createLi(obj);
			var p = this.getDock();
			p.appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'BR') {
			o = this.createBr(obj);
			var p = this.getDock();
			if (p.tagName != 'text:p') { // td
				var p0 = this.createP(obj.parentNode);
				this.docks.push(p0);
				p0.appendChild(o);
			} else {
				p.appendChild(o);
			}
			this.lastWR = null;
		} else if (t == 'IMG') {
			var p = this.getDock();
			var o = this.createPic(obj);
			if (p.tagName != 'text:p') { // td
				var p0 = this.createP(obj.parentNode);
				p0.appendChild(o);
				p.appendChild(p0);
				this.docks.push(p0);
				endPop = true;
			} else {
				p.appendChild(o);
			}
			this.lastWR = null;
		} else if (t == 'TABLE') {
			var prt = this.getDockTable();
			o = this.createTable(obj);
			endPop = true;
			prt.appendChild(o);
			this.docks.push(o);
		} else if (t == 'TBODY') {

		} else if (t == 'TR') {
			o = this.createTr(obj);
			this.getDock().appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'TD') {
			o = this.createTd(obj);
			this.getDock().appendChild(o);
			if (o.tagName == 'table:covered-table-cell') {
				return;
			}
			this.docks.push(o);

			endPop = true;
		} else if (t == 'HR') {
			var p = this.getDockTable();
			o = this.createP(obj);
			var wr = this.createSpan(obj);
			o.appendChild(wr);
			var hr = this.createHr();
			wr.appendChild(hr);
			this.lastWR = null;
			p.appendChild(o);
		} else if (t == 'H1' || t == 'H2' || t == 'H3' || t == 'P'
				|| t == 'CENTER' || t == 'DIV') {
			var p = this.getDockTable();
			o = this.createP(obj);

			p.appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'SCRIPT') {
			return;
		} else if (t == 'BODY' || t == 'OL' || t == 'UL') {

		} else {

		}
		for (var i = 0; i < obj.childNodes.length; i++) {
			var ochild = obj.childNodes[i];
			this.walker1(ochild);
		}
		if (endPop) {
			this.docksPop(o);
			this.lastWR = null;
		}
	};
	this.walkerTxt = function(obj) {
		if (obj.nodeValue.trimEx() == "") {
			this.lastWR = null;
			return;
		}
		var eleTxt = this.createSpan(obj.parentNode);
		var v = obj.nodeValue.replace(/ /ig, ' ');
		if (obj.parentNode.tagName == 'TD' || obj.parentNode.tagName == 'P') {
			v = v.trim();
		}
		if (EWA.B.IE) {
			eleTxt.text = v;
		} else {
			eleTxt.textContent = v;
		}
		var p = this.getDock();
		if (p.tagName != 'text:p') { // td
			var p0 = this.createP(obj.parentNode);
			p0.appendChild(eleTxt);
			p.appendChild(p0);
			this.docks.push(p0);
		} else {
			p.appendChild(eleTxt);
		}

	};
	this.fixTable = function() {
		var objs = $('td[colspan]').toArray();
		for (var m = 0; m < objs.length; m++) {
			var o = objs[m];
			var tr = o.parentNode;
			var a = o.colSpan;
			o.setAttribute('m_colspan', a);
			for (var i = 0; i < a - 1; i++) {
				var td = tr.insertCell(o.cellIndex + i + 1);
				td.innerHTML = i;
				td.bgColor = 'blue';
				td.className = 'colspan';
				td.setAttribute('fixed', '1');
				td.style.display = 'none';
			}
			// o.colSpan = "";
		}

		$('td[rowspan]').each(
				function() {
					var a = this.rowSpan;
					var tr = this.parentNode;
					var tb = tr.parentNode.parentNode;
					for (var i = 0; i < a - 1; i++) {
						var td = tb.rows[tr.rowIndex + i + 1]
								.insertCell(this.cellIndex);
						td.innerHTML = i;
						td.bgColor = 'green';
						td.className = 'rowspan';
						td.setAttribute('fixed', '1');
					}
					$(this).attr('m_rowspan', a);
					this.rowSpan = "";
				});

	};
	this.unFixTable = function() {
		$('td[fixed]').each(function() {
			this.parentNode.removeChild(this);
		});
		$('td[m_colspan]').each(function() {
			this.colSpan = this.getAttribute('m_colspan');
		});
		$('td[m_rowspan]').each(function() {
			this.rowSpan = this.getAttribute('m_rowspan');
		});
	}
	this.createEle = function(tag) {
		var ele8 = this.doc.XmlDoc.createElement(tag);
		return ele8;
	};
	this.createEles = function(tags) {
		var tt = tags.split(',');
		var rts = [];
		for (var i = 0; i < tt.length; i++) {
			var ele8 = this.createEle(tt[i].trim());
			rts.push(ele8);
			if (i > 0) {
				rts[0].appendChild(ele8);
			}
		}
		return rts;
	};
	this.createElesLvl = function(tags) {
		var tt = tags.split(',');
		var rts = [];
		for (var i = 0; i < tt.length; i++) {
			var ele8 = this.createEle(tt[i].trim());
			rts.push(ele8);
			if (i > 0) {
				rts[i - 1].appendChild(ele8);
			}
		}
		return rts;
	};
	this.createElesSameLvl = function(tags, pNode) {
		var tt = tags.split(',');
		var rts = [];
		for (var i = 0; i < tt.length; i++) {
			var ele8 = this.createEle(tt[i].trim());
			rts.push(ele8);
			pNode.appendChild(ele8);
		}
		return rts;
	}

	this.createHr = function() {
		// <w:pict w14:anchorId="0C1134BF">
		// <v:rect id="_x0000_i1037" style="width:.05pt;height:1pt"
		// o:hralign="center" o:hrstd="t"
		// o:hrnoshade="t" o:hr="t" fillcolor="black [3213]" stroked="f"/>
		// </w:pict>
		var hr = this.createElesLvl('w:pict,v:rect');
		hr[0].setAttribute('w14:anchorId', this.getParaId());
		this.setAtts(hr[1], {
			style : "width:.05pt;height:1pt",
			"o:hralign" : "center",
			"o:hrstd" : "t",
			"o:hrnoshade" : "t",
			"o:hr" : "t",
			fillcolor : "black [3213]",
			stroked : "f"
		});
		return hr[0];
	};
	this.createUl = function(obj) {
		// <text:list xml:id="list6985372980825310444" text:style-name="L1">
		var ul = this.createEle('text:list');
		var stName = obj.tagName == 'OL' ? 'L1' : 'L2';
		ul.setAttribute('text:style-name', stName);
		return ul;
	}
	this.createLi = function(obj) {
		var li = this.createEle('text:list-item');
		return li;
	};
	this.createStyle = function(obj) {
		// <style:style style:name="P5" style:family="paragraph"
		// style:parent-style-name="Standard">
		// <style:paragraph-properties fo:text-align="end"
		// style:justify-single-word="false"/>
		// <style:text-properties style:font-name="黑体" fo:font-size="16pt"
		// style:font-name-asian="黑体" style:font-size-asian="16pt"
		// style:font-size-complex="16pt"/>
		// </style:style>
		var st = this.createEle("style:style");
		st.setAttribute('style:family', "text");
		var o = $(obj);

		// font family && font size
		var tp = this.createEle("style:text-properties");
		st.appendChild(tp);

		if (o.css('font-family') != '') {// color
			var f = o.css('font-family').replace(/\'/ig, "").split(',');

			var exp = /[a-z]/ig;
			var f1 = (f[0].trim().toUpperCase() == 'Microsoft YaHei'
					.toUpperCase()) ? 'Microsoft YaHei'
					: ((exp.test(f[0])) ? '宋体' : f[0].trim());
			if (f.length == 1) {
				var f2 = 'Arial';
			} else {
				var f2 = (exp.test(f[1])) ? f[1].trim() : 'Arial';
			}

			var fs = this.fontSize(o.css('font-size'));
			var css = {
				"style:font-name" : f2,
				"style:font-name-asian" : f1,
				"fo:font-size" : fs,
				"style:font-size-complex" : fs,
				"style:font-size-asian" : fs
			}
			if (f.indexOf('Microsoft YaHei') >= 0) {
				// console.log(css)
			}
			this.setAtts(tp, css);
		}

		if (o.css('color') != '') {// color
			var c1 = '#' + this.rgb1(o.css('color')).toLowerCase();
			tp.setAttribute("fo:color", c1);
		}
		var b = o.css('font-weight');
		if (o.tagName == 'B' || !(b == '' || b == 'normal' || b == '400')) {
			tp.setAttribute("fo:font-weight", "bold");
			tp.setAttribute("style:font-weight-complex", "bold");
			tp.setAttribute("style:font-weight-asian", "bold");
		}
		if (o.tagName == 'I') {
			tp.setAttribute("fo:font-style", "italic");
			tp.setAttribute("style:font-style-complex", "italic");
			tp.setAttribute("style:font-style-asian", "italic")
		}
		return this.checkExistsStyle(st);
	};
	this.createStyleBreak = function( ) {
		/*
		 * 	<style:style style:name="P1" style:family="paragraph"
			style:parent-style-name="Standard">
			<style:paragraph-properties
				fo:break-before="page" />
		</style:style>
		 */
		var st = this.createEle("style:style");
		st.setAttribute('style:family', "paragraph");
		var stpp = this.createEle("style:paragraph-properties");
		st.appendChild(stpp);
		this.setAtts(stpp, {
			"fo:break-before" : "page"
		});
		return this.checkExistsStyle(st);
	};
	this.createStyleP = function(obj) {
		// <style:style style:name="P5" style:family="paragraph"
		// style:parent-style-name="Standard">
		// <style:paragraph-properties fo:text-align="end"
		// style:justify-single-word="false" fo:margin-left="1.482cm"
		// fo:margin-right="0cm" fo:line-height="150%"/>
		// </style:style>
		var st = this.createEle("style:style");
		st.setAttribute('style:family', "paragraph");
		var stpp = this.createEle("style:paragraph-properties");
		st.appendChild(stpp);
		var o = $(obj);
		var al = o.css('text-align');
		al = al == null ? "" : al;
		if (al.indexOf('center') >= 0) {
			al = "center";// 左对齐
		} else if (al.indexOf('right') >= 0) {
			al = "end";// 左对齐
		} else {
			al = "";
		}
		// align
		if (al != "") {
			this.setAtts(stpp, {
				"fo:text-align" : al
			});
		}
		var ml = o.css('margin-left').replace('px', '');
		var mr = o.css('margin-right').replace('px', '');

		var mlcm = this.px2cm(ml) + 'cm';
		var mrcm = this.px2cm(mr) + 'cm';

		this.setAtts(stpp, {
			"fo:margin-left" : mlcm,
			"fo:margin-right" : mrcm,
			"fo:text-indent" : "0cm",
			"style:justify-single-word" : "false"
		});
		// var mt = o.css('margin-top').replace('px', '');
		// var mb = o.css('margin-bottom').replace('px', '');
		// var lh = o.css('font-size').replace('px', '');
		// var lh1 = ((mt * 1.0 + mb * 1 + lh * 1) / lh -1)* 100;
		// this.setAtts(stpp, {
		// "fo:margin-left" : mlcm,
		// "fo:margin-right" : mrcm,
		// "fo:line-height" : lh1 + "%"
		// });
		return this.checkExistsStyle(st);
	};
	/**
	 * 字体是否已经存在
	 * 
	 * @param {Object}
	 *            st
	 * @memberOf {TypeName}
	 * @return {TypeName}
	 */
	this.checkExistsStyle = function(st) {
		var s = st.outerHTML;
		if (this.fontMap[s]) {
			return this.fontMap[s];
		}
		var stName = this.getStyleName();
		this.setAtts(st, {
			"style:name" : stName
		});
		this.fontMap[s] = stName;
		this.doc.XmlDoc.getElementsByTagName("automatic-styles")[0]
				.appendChild(st);
		// console.log(st)
		return stName;
	}
	this.createStyleTd = function(obj) {
		// <style:style style:name="表格1.A1" style:family="table-cell"
		// style:data-style-name="N0">
		// <style:table-cell-properties fo:padding="0.097cm"
		// fo:border-left="0.002cm solid #000000"
		// style:vertical-align="middle" fo:background-color="transparent"
		// fo:border-right="none" fo:border-top="0.002cm solid #000000"
		// fo:border-bottom="0.002cm solid #000000"/>
		// </style:style>
		var st = this.createEle("style:style");
		this.setAtts(st, {
			"style:family" : "table-cell",
			"style:data-style-name" : "NO"
		});
		var stpp = this.createEle("style:table-cell-properties");
		st.appendChild(stpp);
		this.setAtts(stpp, {
			"fo:background-color" : "transparent",
			"fo:padding" : "0.097cm"
		});
		for (var i = 0; i < this.EWA_DocTmp.borders.length; i++) {
			var n = this.EWA_DocTmp.borders[i];
			var b1 = this.createBorder(obj, n);
			stpp.setAttribute("fo:border-" + n, b1);
		}
		var o = $(obj);
		var vAlign = o.css("vertical-align");
		if (vAlign == 'middle') {
			stpp.setAttribute('style:vertical-align', 'middle');
		} else if (vAlign == 'bottom') {
			stpp.setAttribute('style:vertical-align', 'bottom');
		}
		return this.checkExistsStyle(st);
	};
	this.createStyleCol = function(obj) {
		// <style:style style:name="表格1.A" style:family="table-column">
		// <style:table-column-properties style:column-width="8.498cm"/>
		// </style:style>
		var st = this.createEle("style:style");
		this.setAtts(st, {
			"style:family" : "table-column"
		});
		var stpp = this.createEle("style:table-column-properties");
		st.appendChild(stpp);
		var w = $(obj).width();
		var w1 = this.px2cm(w);
		stpp.setAttribute('style:column-width', w1 + 'cm');
		var st1 = this.checkExistsStyle(st);
		// console.log(st)
		return st1;
	};
	this.createStyleTable = function(obj) {
		//
		// <style:style style:name="表格1" style:family="table">
		// <style:table-properties style:width="17cm" table:align="center"
		// style:shadow="none"/>
		// </style:style>
		var st = this.createEle("style:style");
		this.setAtts(st, {
			"style:family" : "table"
		});
		var stpp = this.createEle("style:table-properties");
		st.appendChild(stpp);
		var w = $(obj).width();
		var w1 = this.px2cm(w);
		stpp.setAttribute('style:width', w1 + 'cm');
		stpp.setAttribute("style:shadow", "none");
		return this.checkExistsStyle(st);
	};

	this.getStyleName = function() {
		this.fontIndex++;
		return "ST" + this.fontIndex;
	};
	this.fontIndex = 0;
	this.fontMap = {};
	this.createSpan = function(obj) {
		// <text:span text:style-name="T5">
		var r = this.createEle("text:span");
		var stName = this.createStyle(obj);
		r.setAttribute("text:style-name", stName);
		return r;
	};
	this.createBr = function(obj) {
		var bele = this.createEle("text:line-break");
		return bele;
	};
	this.createTable = function(obj) {
		var bele = this.createEle("table:table");
		// <table:table table:name="表格1" table:style-name="表格1">
		bele.setAttribute('table:name', 'gdx' + Math.random());
		var maxCellsTr = {
			row : null,
			num : -1
		};
		if (obj.id == 'EWA_FRAME_G1375752461') {
			var zzzzzzzz = 1;
		}
		for (var i = 0; i < obj.rows.length; i++) {
			var tr = obj.rows[i];
			var num = 0;
			for (var m = 0; m < tr.cells.length; m++) {
				var td = tr.cells[m];
				if (td.getAttribute('fixed')) { // 生成前补充的td
					continue;
				}
				num++;
			}
			if (num > maxCellsTr.num) {
				maxCellsTr.num = num;
				maxCellsTr.row = tr;
			}
		}
		// console.log(maxCellsTr)
		if (maxCellsTr.row) {
			for (var i = 0; i < maxCellsTr.row.cells.length; i++) {
				var col = this.createEle("table:table-column");
				var colSt = this.createStyleCol(maxCellsTr.row.cells[i]);
				col.setAttribute('table:style-name', colSt);
				bele.appendChild(col);
			}
			// console.log(bele)
		}
		var tbSt = this.createStyleTable(obj);
		bele.setAttribute('table:style-name', tbSt);
		return bele;
	};

	this.createTr = function(obj) {
		// table:table-row
		var tr = this.createEle("table:table-row");
		return tr;
	};
	this.paraId = 0;
	this.getParaId = function() {
		var v = "0000" + this.paraId;
		v = v.substring(v.length - 4);
		this.paraId++;
		return "F0C0" + v;
	}
	this.createTd = function(obj) {
		// <table:table-cell table:style-name="A2" office:value-type="string">
		if (obj.className == 'rowspan' || obj.className == 'colspan') {
			return this.createEle('table:covered-table-cell');
		}
		var bele = this.createEle("table:table-cell");
		if (obj.getAttribute("m_rowspan") > 1) {
			bele.setAttribute('table:number-rows-spanned', obj
					.getAttribute("m_rowspan"));
		}
		if (obj.getAttribute("m_colspan") > 1) {
			bele.setAttribute('table:number-columns-spanned', obj
					.getAttribute("m_colspan"));
		}
		bele.setAttribute('office:value-type', "string");

		var stName = this.createStyleTd(obj);
		bele.setAttribute('table:style-name', stName);
		return bele;
	};
	this.createWidth = function(obj, tag) {
		// <w:tblW w:w="8702" w:type="dxa" />
		var e = this.createEle('w:' + tag);
		var w = $(obj).width() * 15; // px-->word width
		e.setAttribute('w:w', w);
		e.setAttribute('w:type', "dxa");
		return e;
	};
	this.createHeight = function(obj, tag) {
		// <w:trHeight w:val="10121"/>
		var e = this.createEle('w:' + tag);
		var h = $(obj).height() * 15; // px-->word width
		e.setAttribute('w:val', h);
		return e;
	};
	/*
	 * <style:table-cell-properties fo:padding="0.097cm" fo:border-left="0.002cm
	 * solid #000000" fo:border-right="none" fo:border-top="0.002cm solid
	 * #000000" fo:border-bottom="0.002cm solid #000000"/>
	 */
	this.createBorder = function(obj, a) {
		// "0.002cm solid #000000"
		var v1 = "";
		var v = $(obj).css('border-' + a + '-width');
		var e = this.createEle('w:' + a);
		if (v == '0px') {
			v1 = "none";
		} else {
			v1 = "0.002cm ";
			v1 += $(obj).css('border-' + a + '-style');
			var c = $(obj).css('border-' + a + '-color');
			var c1 = this.rgb1(c).toLowerCase();
			v1 += " #" + c1;
		}
		return v1;
	};
	/**
	 * 不可见的分割，用于两个紧连的表分割等
	 * 
	 * @param {Object}
	 *            obj
	 * @memberOf {TypeName}
	 * @return {TypeName}
	 */
	this.createPVanish = function() {
		var elep = this.createElesLvl("w:p,w:pPr,w:rPr,w:vanish");
		return elep;
	};
	this.createP = function(obj) {
		// <text:p text:style-name="Standard"/>
		var elep = this.createEle("text:p");

		var stName = this.createStyleP(obj);
		elep.setAttribute("text:style-name", stName);
		return elep;
	};
	/**
	 * 创建分页
	 */
	this.createBreak = function() {
		var elep = this.createEle("text:p");
		var stName=this.createStyleBreak();
		elep.setAttribute("text:style-name", stName);
		return elep;
	};
	this.createPic = function(obj) {
		/*
		 * <draw:frame draw:style-name="fr1" draw:name="图形1"
		 * text:anchor-type="as-char" svg:width="4.233cm" svg:height="4.233cm"
		 * draw:z-index="0"> <draw:image
		 * xlink:href="Pictures/10000201000000A0000000A0214A9447.png"
		 * xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>
		 * </draw:frame>
		 */
		var o = $(obj);
		var picName = obj.src;
		this.pics.push(obj.src);

		var idx = picName.lastIndexOf('/');
		picName = picName.substring(idx + 1);
		var emuX = this.px2cm(o.width());
		var emuY = this.px2cm(o.height());

		var rs = this.createElesLvl("draw:frame,draw:image");
		this.setAtts(rs[0], {
			'svg:width' : emuX + 'cm',
			'svg:height' : emuY + 'cm',
			"draw:z-index" : "0",
			"draw:style-name" : "fr1",
			"text:anchor-type" : "as-char"
		});
		this.setAtts(rs[1], {
			'xlink:type' : 'simple',
			'xlink:show' : 'embed',
			"xlink:actuate" : "onLoad",
			"xlink:href" : "{[PIC" + (this.pics.length - 1) + "]}"
		});

		return rs[0];
	};
	this.createEleNs = function(tag, ns) {
		var ele8 = this.doc.XmlDoc.createElementNS(ns, tag);
		return ele8;
	}
	this.setAtts = function(node, params) {
		for ( var n in params) {
			node.setAttribute(n, params[n]);
		}
	}
	this.fontSize = function(f) {
		if (f.indexOf('px') > 0) {
			var f0 = this.EWA_DocTmp.f[f];
			if (f0 == null) {
				f = '9pt';
			} else {
				f = f0;
			}
		}
		var f1 = f.replace('pt', '');
		return f1;
	};
	this.rgb1 = function(s1) {
		var s = s1.replace('rgb(', '').replace(')', '');
		var ss = s.split(',');
		return this.rgb(ss[0] * 1, ss[1] * 1, ss[2] * 1).toUpperCase();
	};
	this.rgb = function(r, g, b) {
		var r1 = r.toString(16);
		var g1 = g.toString(16);
		var b1 = b.toString(16);
		return (r1.length < 2 ? "0" : "") + r1 + (g1.length < 2 ? "0" : "")
				+ g1 + (b1.length < 2 ? "0" : "") + b1;
	};

	this.init = function() {
		this.doc = new EWA_XmlClass();
		this.doc.LoadXml(this.EWA_DocTmp.document);
		if (EWA.B.IE) {
			this.docks
					.push(this.doc.XmlDoc.getElementsByTagName('office:text')[0]);
		} else {
			this.docks.push(this.doc.XmlDoc.getElementsByTagName('text')[0]);
		}
	};
	this.getDock = function() {
		return this.docks[this.docks.length - 1];
	};
	this.idx = 0
	this.docksPop = function(o) {
		while (1 == 1) {
			if (this.docks.length == 1) {
				break;
			}
			var o1 = this.docks.pop();
			if (o1 == o) {
				// console.log(this.docks);
				break;
			}
		}
	};
	this.getDockTable = function() {
		for (var i = this.docks.length - 1; i >= 0; i--) {
			var o = this.docks[i];
			if (o.tagName == 'office:text' || o.tagName == 'table:table-cell'
					|| o.tagName == 'text:list-item') {
				return o;
			}
		}
	};
	/**
	 * 包括英制单位（914,400 个 EMU 单位为 1 英寸）
	 * 
	 * @param {Object}
	 *            v
	 * @return {TypeName}
	 */
	this.px2cm = function(v) {
		if (v == null || v == '') {
			return 0;
		}
		return v / 96 * 2.539999918;
	}
	this.EWA_DocTmp = {
		document : '<?xml version="1.0" encoding="UTF-8"?>\
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"\
                         xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"\
                         xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"\
                         xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"\
                         xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"\
                         xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"\
                         xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/"\
                         xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"\
                         xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"\
                         xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"\
                         xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"\
                         xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"\
                         xmlns:math="http://www.w3.org/1998/Math/MathML"\
                         xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"\
                         xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"\
                         xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer"\
                         xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events"\
                         xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema"\
                         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\
                         xmlns:rpt="http://openoffice.org/2005/report"\
                         xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"\
                         xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#"\
                         xmlns:tableooo="http://openoffice.org/2009/table"\
                         xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"\
                         office:version="1.2"><office:scripts/>\
    <office:font-face-decls>\
        <style:font-face style:name="OpenSymbol" svg:font-family="OpenSymbol"/>\
        <style:font-face style:name="微软雅黑" svg:font-family="微软雅黑"/>\
        <style:font-face style:name="Lucida Sans1" svg:font-family="&apos;Lucida Sans&apos;"\
                         style:font-family-generic="swiss"/><style:font-face style:name="宋体" svg:font-family="宋体" style:font-pitch="variable"/>\
        <style:font-face style:name="黑体" svg:font-family="黑体" style:font-pitch="variable"/>\
        <style:font-face style:name="Times New Roman" svg:font-family="&apos;Times New Roman&apos;"\
                         style:font-family-generic="roman" style:font-pitch="variable"/>\
        <style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss"\
                         style:font-pitch="variable"/>\
        <style:font-face style:name="Lucida Sans" svg:font-family="&apos;Lucida Sans&apos;"\
                         style:font-family-generic="system" style:font-pitch="variable"/>\
        <style:font-face style:name="微软雅黑1" svg:font-family="微软雅黑" style:font-family-generic="system"\
                         style:font-pitch="variable"/>\
    </office:font-face-decls>\
    <office:automatic-styles><style:style style:name="fr1" style:family="graphic" style:parent-style-name="Graphics">\
            <style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline"\
                                      style:horizontal-pos="center" style:horizontal-rel="paragraph" style:shadow="none"\
                                      style:mirror="none" fo:clip="rect(0cm, 0cm, 0cm, 0cm)" draw:luminance="0%"\
                                      draw:contrast="0%" draw:red="0%" draw:green="0%" draw:blue="0%" draw:gamma="100%"\
                                      draw:color-inversion="false" draw:image-opacity="100%"\
                                      draw:color-mode="standard"/></style:style><text:list-style style:name="L1">\
            <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="1.27cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="1.905cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="2.54cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="4" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="3.175cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="3.175cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="5" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="3.81cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="3.81cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="6" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="4.445cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="4.445cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="7" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="5.08cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="5.08cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="8" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="5.715cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="5.715cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="9" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="6.35cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="6.35cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
            <text:list-level-style-number text:level="10" text:style-name="Numbering_20_Symbols" style:num-suffix="."\
                                          style:num-format="1">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="6.985cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="6.985cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-number>\
        </text:list-style>\
        <text:list-style style:name="L2">\
            <text:list-level-style-bullet text:level="1" text:style-name="Bullet_20_Symbols" text:bullet-char="•">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="1.27cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="2" text:style-name="Bullet_20_Symbols" text:bullet-char="◦">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="1.905cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="3" text:style-name="Bullet_20_Symbols" text:bullet-char="▪">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="2.54cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="4" text:style-name="Bullet_20_Symbols" text:bullet-char="•">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="3.175cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="3.175cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="5" text:style-name="Bullet_20_Symbols" text:bullet-char="◦">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="3.81cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="3.81cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="6" text:style-name="Bullet_20_Symbols" text:bullet-char="▪">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="4.445cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="4.445cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="7" text:style-name="Bullet_20_Symbols" text:bullet-char="•">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="5.08cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="5.08cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="8" text:style-name="Bullet_20_Symbols" text:bullet-char="◦">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="5.715cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="5.715cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="9" text:style-name="Bullet_20_Symbols" text:bullet-char="▪">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="6.35cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="6.35cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
            <text:list-level-style-bullet text:level="10" text:style-name="Bullet_20_Symbols" text:bullet-char="•">\
                <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">\
                    <style:list-level-label-alignment text:label-followed-by="listtab"\
                                                      text:list-tab-stop-position="6.985cm" fo:text-indent="-0.635cm"\
                                                      fo:margin-left="6.985cm"/>\
                </style:list-level-properties>\
            </text:list-level-style-bullet>\
        </text:list-style>\
        <number:number-style style:name="N0">\
            <number:number number:min-integer-digits="1"/>\
        </number:number-style></office:automatic-styles><office:body><office:text /></office:body></office:document-content>',
		f : {
			"9px" : "7pt",
			"10px" : "7.5pt",
			"11px" : "8.5pt",
			"12px" : "9pt",
			"13px" : "10pt",
			"14px" : "10.5pt",
			"14.8px" : "10.5pt",
			"15px" : "11.5pt",
			"16px" : "12pt",
			"17px" : "13pt",
			"18px" : "13.5pt",
			"19px" : "14.5pt",
			"20px" : "15pt",
			"21px" : "16pt",
			"22px" : "16.5pt",
			"23px" : "17.5pt",
			"24px" : "18pt",
			"25px" : "19pt",
			"26px" : "19.5pt",
			"27px" : "20.5pt",
			"28px" : "21pt",
			"29px" : "22pt",
			"30px" : "22.5pt",
			"31px" : "23.5pt",
			"32px" : "24pt",
			"33px" : "25pt",
			"34px" : "25.5pt",
			"35px" : "26.5pt",
			"36px" : "27pt",
			"37px" : "28pt",
			"38px" : "28.5pt",
			"39px" : "29.5pt",
			"40px" : "30pt",
			"41px" : "31pt",
			"42px" : "31.5pt",
			"43px" : "32.5pt",
			"44px" : "33pt",
			"45px" : "34pt",
			"46px" : "34.5pt",
			"47px" : "35.5pt",
			"48px" : "36pt",
			"49px" : "37pt",
			"50px" : "37.5pt",
			"51px" : "38.5pt",
			"52px" : "39pt",
			"53px" : "40pt",
			"54px" : "40.5pt",
			"55px" : "41.5pt",
			"56px" : "42pt",
			"57px" : "43pt",
			"58px" : "43.5pt",
			"59px" : "44.5pt",
			"60px" : "45pt",
			"61px" : "46pt",
			"62px" : "46.5pt",
			"63px" : "47.5pt",
			"64px" : "48pt",
			"65px" : "49pt",
			"66px" : "49.5pt",
			"67px" : "50.5pt",
			"68px" : "51pt",
			"69px" : "52pt",
			"70px" : "52.5pt",
			"71px" : "53.5pt",
			"72px" : "54pt",
			"73px" : "55pt",
			"74px" : "55.5pt",
			"75px" : "56.5pt",
			"76px" : "57pt",
			"77px" : "58pt",
			"78px" : "58.5pt",
			"79px" : "59.5pt",
			"80px" : "60pt",
			"81px" : "61pt",
			"82px" : "61.5pt",
			"83px" : "62.5pt",
			"84px" : "63pt",
			"85px" : "64pt",
			"86px" : "64.5pt",
			"87px" : "65.5pt",
			"88px" : "66pt",
			"89px" : "67pt",
			"90px" : "67.5pt",
			"91px" : "68.5pt",
			"92px" : "69pt",
			"93px" : "70pt",
			"94px" : "70.5pt",
			"95px" : "71.5pt",
			"96px" : "72pt",
			"97px" : "73pt",
			"98px" : "73.5pt",
			"99px" : "74.5pt"
		},
		borders : [ 'top', 'bottom', 'right', 'left' ]
	};
	this.postData = function() {
		var data = {
			xml : this.xml(),
			pics : this.pics.join(','),
			ols : this.liNum
		};
		return data;
	}
	this.init();
	this.xml = function() {
		return this.doc.GetXml();
	};
};
// var word = new EWA_DocWordClass();
// word.walker(oo);
// word.xml();
