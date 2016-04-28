function EWA_DocWordClass() {
	this.docks = [];
	this.pics = [];
	this.liNum = 0;//numering.xml defined
	this.walker = function(obj, pNode) {
		if (obj.nodeType == 3) {
			this.walkerTxt(obj)
		} else if (obj.nodeType == 1) {
			this.walkerEle(obj)
		}
	};
	this.walkerEle = function(obj) {
		if(obj.style.display=='none'){
			return;
		}
		var t = obj.tagName;
		var endPop = false;
		var o;
		if (t == 'LI') {
			o = this.createP(obj);
			var p = this.getDockTable();
			p.appendChild(o);
			this.docks.push(o);
			endPop = true;
			this.createLi(obj);
		} else if (t == 'BR') {
			o = this.createBr(obj);
			var p = this.getDock();
			//var o = this.createSpan(obj);
			if (p.tagName != 'w:p') { //td
				var p0 = this.createP(obj.parentNode);
				var wr = this.createSpan(obj.parentNode);
				p0.appendChild(wr);
				p0.appendChild(o);
				p.appendChild(p0);
				this.docks.push(p0);

			} else {
				p.appendChild(o);
			}
			this.lastWR = null;
		} else if (t == 'IMG') {
			var p = this.getDock();
			var o = this.createPic(obj);
			if (p.tagName != 'w:p') { //td
				var p0 = this.createP(obj.parentNode);
				var wr = this.createSpan(obj.parentNode);
				p0.appendChild(wr);
				p0.appendChild(o);
				p.appendChild(p0);
				this.docks.push(p0);
			} else {
				p.appendChild(o);
			}
			this.lastWR = null;
		} else if (t == 'TABLE') {
			var prt = this.getDockTable();
			var prev = obj.previousElementSibling;
			if (prev != null && prev.tagName == "TABLE") {
				var p = this.createPVanish();
				prt.appendChild(p[0]);
			}
			o = this.createTable(obj);
			//jzp
			if(o!=null){
				endPop = true;
				prt.appendChild(o);
				this.docks.push(o);
			}
		} else if (t == 'TBODY') {
			o = this.createTbody(obj);
			this.getDock().appendChild(o);
		} else if (t == 'TR') {
			o = this.createTr(obj);
			this.getDock().appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'TD') {
			o = this.createTd(obj);
			this.getDock().appendChild(o);
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
		} else if (obj.parentNode.tagName != 'LI'
				&& (t == 'H1' || t == 'H2' || t == 'H3' || t == 'P'
						|| t == 'CENTER' || t == 'DIV')) {
			var p = this.getDockTable();
			o = this.createP(obj);

			p.appendChild(o);
			this.docks.push(o);
			endPop = true;
		} else if (t == 'SCRIPT') {
			return;
		} else if (t == 'BODY' || t == 'OL' || t == 'UL') {

		} else {
			var p = this.getDock();
			o = this.createSpan(obj);
			if (p.tagName != 'w:p') { //td
				var p0 = this.createP(obj.parentNode);
				p0.appendChild(o);
				p.appendChild(p0);
				this.docks.push(p0);
			} else {
				p.appendChild(o);
			}
		}
		for ( var i = 0; i < obj.childNodes.length; i++) {
			var ochild = obj.childNodes[i];
			this.walker(ochild);
		}
		if (endPop) {
			this.docksPop(o);
			this.lastWR = null;
		}
		if (t == 'TD' && o.childNodes.length == 1) {
			var p0 = this.createEle('w:p')
			o.appendChild(p0);
		} else if (t == 'TABLE' && o.nextSibling == null
				&& o.parentNode.tagName == 'w:tc') {
			var p0 = this.createEle('w:p')
			o.parentNode.appendChild(p0);
		}
	}
	this.walkerTxt = function(obj) {
		if (obj.nodeValue.trimEx() == "") {
			this.lastWR = null;
			return;
		}
		var eleTxt = this.createText(obj);
		var p = this.getDock();
		if (p.tagName != 'w:p') { //td
			var p0 = this.createP(obj.parentNode);
			var wr = this.createSpan(obj.parentNode);
			p0.appendChild(wr);
			wr.appendChild(eleTxt);

			p.appendChild(p0);
			this.docks.push(p0);
		} else {
			if (!this.lastWR) {
				var wr = this.createSpan(obj.parentNode);
				p.appendChild(wr)
			}
			this.lastWR.parentNode.appendChild(eleTxt);
		}
		this.lastWR = null;
	}
	this.createEle = function(tag) {
		var ele8 = this.doc.XmlDoc.createElement(tag);
		return ele8;
	};
	this.createEles = function(tags) {
		var tt = tags.split(',');
		var rts = [];
		for ( var i = 0; i < tt.length; i++) {
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
		for ( var i = 0; i < tt.length; i++) {
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
		for ( var i = 0; i < tt.length; i++) {
			var ele8 = this.createEle(tt[i].trim());
			rts.push(ele8);
			pNode.appendChild(ele8);
		}
		return rts;
	}
	this.createText = function(obj) {
		if (obj.nodeValue.trim() == "") {
			return;
		}
		var t = this.createEle("w:t");
		if (EWA.B.IE) {
			t.text = obj.nodeValue;
		} else {
			t.textContent = obj.nodeValue;
		}
		return t;
	};
	this.createHr = function() {
		//<w:pict w14:anchorId="0C1134BF">
		//   <v:rect id="_x0000_i1037" style="width:.05pt;height:1pt" o:hralign="center" o:hrstd="t"
		//        o:hrnoshade="t" o:hr="t" fillcolor="black [3213]" stroked="f"/>
		//</w:pict>
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
	}
	this.createLi = function(obj) {
		// <w:numPr> <w:ilvl w:val="0"/> <w:numId w:val="8"/></w:numPr>
		//<w:spacing w:line="270" w:lineRule="atLeast"/>
		var numPrs = this.createEles('w:numPr,w:ilvl,w:numId');
		if (obj.parentNode.tagName == 'OL') {
			if (obj == obj.parentNode.getElementsByTagName('li')[0]) {
				this.liNum++;
			}
			numPrs[1].setAttribute('w:val', 0);
			numPrs[2].setAttribute('w:val', this.liNum);
		} else {
			numPrs[1].setAttribute('w:val', 0);
			numPrs[2].setAttribute('w:val', 22);
		}
		var wSpace = this.createEle('w:spacing');
		wSpace.setAttribute('w:line', 270);
		wSpace.setAttribute('w:lineRule', "atLeast");

		var p = this.getDock();
		p.getElementsByTagName('w:pPr')[0].appendChild(numPrs[0]);
		p.getElementsByTagName('w:pPr')[0].appendChild(wSpace);
	}
	this.createSpan = function(obj) {
		var r = this.createEle("w:r");
		var o = $(obj);
		var wrpr = this.createEle("w:rPr");

		this.lastWR = wrpr;

		r.appendChild(wrpr);
		if (o.css('color') != '') {// color
			var c = this.createEle('w:color');
			var c1 = this.rgb1(o.css('color'));
			c.setAttribute("w:val", c1);
			wrpr.appendChild(c);
		}
		if (o.css('font-family') != '') {// color
			// <w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="Times New Roman"  w:hint="eastAsia"/>
			var f = o.css('font-family').replace(/\'/ig, "").split(',');
			var c = this.createEle('w:rFonts');
			var exp = /[a-z]/ig;
			var f1 = (exp.test(f[0])) ? '宋体' : f[0].trim();
			var f2 = (exp.test(f[0])) ? f[0].trim() : 'Times New Roman';
			this.setAtts(c, {
				"w:ascii" : f1,
				"w:eastAsia" : f1,
				"w:hAnsi" : f1,
				"w:hint" : "eastAsia",
				"w:cs" : f2
			});
			wrpr.appendChild(c);
		}
		if (o.css('font-size') != '') {// color
			var f = o.css('font-size');
			var f1 = this.fontSize(f);// <w:sz w:val="36" /><w:szCs w:val="36"
			// />
			console.log(f1)
			if (f1 != null) {
				var c = this.createEle('w:sz');
				c.setAttribute("w:val", f1 * 2);
				wrpr.appendChild(c);

				var c1 = this.createEle('w:szCs');
				c1.setAttribute("w:val", f1 * 2);
				wrpr.appendChild(c1);
			}
		}
		var b = o.css('font-weight');
		if (o.tagName == 'B' || !(b == '' || b == 'normal' || b == '400')) {
			var c = this.createEle('w:b');
			wrpr.appendChild(c);
		}
		if (o.tagName == 'I') {
			var c = this.createEle('w:i');
			wrpr.appendChild(c);
		}

		return r;

	};
	this.createBr = function(obj) {
		var bele = this.createEle("w:r");
		var bele1 = this.createEle("w:br");
		bele.appendChild(bele1);

		return bele;
	};
	this.createTable = function(obj) {
		var bele = this.createEle("w:tbl");

		var tblPr = this.createEle("w:tblPr");
		var tblStyle = this.createEle("w:tblStyle");
		tblStyle.setAttribute("w:val", "a4");

		tblPr.appendChild(tblStyle);
		//<w:tblW w:w="8702" w:type="dxa" />
		var tbW = this.createWidth(obj, 'tblW');
		tblPr.appendChild(tbW);
		//		if (this.getDockTable().tagName == 'w:body') {
		//			tbW.setAttribute('w:type', 'pct'); //100%
		//			tbW.setAttribute('w:w', 5000);
		//			//console.log(tbW)
		//		}
		obj.setAttribute('ww', tbW.getAttribute('w:w'));
		bele.appendChild(tblPr);

		return bele;
	};
	this.createTbody = function() {
		var bele = this.createEle("w:tblGrid");
		return bele;
	};
	this.createTr = function(obj) {
		var trs = this.createElesLvl("w:tr,w:trPr");
		this.setAtts(trs[0], {
			"w:rsidR" : "003D2D54",
			"w14:textId" : "77777777",
			"w14:paraId" : this.getParaId()
		});
		var h = this.createHeight(obj, 'trHeight');
		trs[1].appendChild(h);
		return trs[0];
	};
	this.paraId = 0;
	this.getParaId = function() {
		var v = "0000" + this.paraId;
		v = v.substring(v.length - 4);
		this.paraId++;
		return "F0C0" + v;
	}
	this.createTd = function(obj) {
		thisTr = obj.parentNode;
		var tb = thisTr.parentNode.parentNode;
		var vm = obj.getAttribute('vmerge');
		var refTdww = 0;
		if (vm != null && vm != '') {
			var bele = this.createEle("w:tc");
			var tcPr = this.createEle("w:tcPr");
			var vMerge = this.createEle("w:vMerge");
			tcPr.appendChild(vMerge);
			bele.appendChild(tcPr);
			this.getDock().appendChild(bele);

			var p = this.createEle('w:p');
			bele.appendChild(p);
			var refIdx = vm.split(',');
			var refTd = tb.rows[refIdx[0]].cells[refIdx[1]];
			tcPr.appendChild(this.createTdBorders(refTd));
			refTdww = refTd.getAttribute('wwtd') * 1;
		}

		var bele = this.createEle("w:tc");
		var tcPr = this.createEle("w:tcPr");
		bele.appendChild(tcPr);

		if (obj.rowSpan > 1) {
			//<w:vMerge w:val="restart" />
			var vMerge = this.createEle("w:vMerge");
			vMerge.setAttribute("w:val", "restart");
			tcPr.appendChild(vMerge);
			for ( var i = 0; i < obj.rowSpan - 1; i++) {
				var tr = tb.rows[i + 1 + thisTr.rowIndex];
				if (tr) {
					var td = tr.cells[obj.cellIndex];
					if (td) {
						td.setAttribute('vmerge', thisTr.rowIndex + ','
								+ obj.cellIndex);
					}
				}

			}
		}
		if (obj.colSpan > 1) {
			// <w:gridSpan w:val="2"/>
			var colSpan = this.createEle('w:gridSpan');
			colSpan.setAttribute('w:val', obj.colSpan);
			tcPr.appendChild(colSpan);
		}
		tcPr.appendChild(this.createTdBorders(obj));

		var tcW = this.createWidth(obj, 'tcW');
		var tr = obj.parentNode;

		var trww = tr.getAttribute('wwtr');
		if (trww == null || trww == '') {
			var tb = tr.parentNode.parentNode;
			var ww = tb.getAttribute('ww');
			trww = ww;
		}
		var w = tcW.getAttribute("w:w");
		if (obj != obj.parentNode.cells[obj.parentNode.cells.length - 1]) {
			//<w:tcW w:w="2901" w:type="dxa" />
			obj.setAttribute('wwtd', w);
			var w1 = trww * 1 - w * 1 - refTdww;
			tr.setAttribute('wwtr', w1);
		} else {
			//最后一个单元格不设置宽度
			//obj.setAttribute('wwtd', trww * 1 - refTdww);
			tcW.setAttribute("w:w", 0);
			tcW.setAttribute("w:type", "auto");
		}
		tcPr.appendChild(tcW);

		var o = $(obj);
		var vAlign = o.css("vertical-align");
		if (vAlign == 'middle') {
			//<w:vAlign w:val="bottom"/>
			var e1 = this.createEle('w:vAlign');
			e1.setAttribute('w:val', 'center');
			tcPr.appendChild(e1);
		} else if (vAlign == 'bottom') {
			var e1 = this.createEle('w:vAlign');
			e1.setAttribute('w:val', 'bottom');
			tcPr.appendChild(e1);
		}
		return bele;
	};
	this.createWidth = function(obj, tag) {
		//<w:tblW w:w="8702" w:type="dxa" />
		var e = this.createEle('w:' + tag);
		var w = $(obj).width() * 15; //px-->word width
		e.setAttribute('w:w', w);
		e.setAttribute('w:type', "dxa");
		return e;
	};
	this.createHeight = function(obj, tag) {
		//<w:trHeight w:val="10121"/>
		var e = this.createEle('w:' + tag);
		var h = $(obj).height() * 15; //px-->word width
		e.setAttribute('w:val', h);
		return e;
	};
	/*
	 * <w:tcBorders>
							<w:top w:val="nil" />
							<w:left w:val="single" w:sz="8" w:space="0" w:color="000000" />
							<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000" />
							<w:right w:val="single" w:sz="8" w:space="0" w:color="000000" />
						</w:tcBorders>
	 * @param {Object} obj
	 */
	this.createTdBorders = function(obj) {
		var e = this.createEle('w:tcBorders');
		for ( var i = 0; i < this.EWA_DocTmp.borders.length; i++) {
			var b1 = this.createBorder(obj, this.EWA_DocTmp.borders[i]);
			e.appendChild(b1);
		}
		return e;
	};
	this.createBorder = function(obj, a) {
		var v = $(obj).css('border-' + a + '-width');
		var e = this.createEle('w:' + a);
		if (v == '0px') {
			e.setAttribute('w:val', 'nil');
		} else {
			e.setAttribute('w:val', 'single');
			var c = $(obj).css('border-' + a + '-color');
			var c1 = this.rgb1(c);
			e.setAttribute('w:color', c1);
			e.setAttribute('w:sz', 6);
		}
		return e;
	};
	/**
	 * 不可见的分割，用于两个紧连的表分割等
	 * @param {Object} obj
	 * @memberOf {TypeName} 
	 * @return {TypeName} 
	 */
	this.createPVanish = function() {
		var elep = this.createElesLvl("w:p,w:pPr,w:rPr,w:vanish");
		return elep;
	};
	this.createP = function(obj) {
		var elep = this.createEle("w:p");
		//<w:pPr>
		//   <w:pStyle w:val="a7"/>
		//   <w:jc w:val="left"/>
		//   <w:rPr>
		//       <w:rFonts w:hint="eastAsia"/>
		//   </w:rPr>
		//</w:pPr>
		//
		var elepPr = this.createEle("w:pPr");
		elep.appendChild(elepPr);

		var pPrs = this.createElesSameLvl("w:pStyle,w:jc", elepPr);
		var elejc = pPrs[1];

		var al = $(obj).css('text-align');
		al = al == null ? "" : al;
		if (al.indexOf('center') >= 0) {
			elejc.setAttribute("w:val", "center");// 左对齐
		} else if (al.indexOf('right') >= 0) {
			elejc.setAttribute("w:val", "right");// 左对齐
		} else {
			elejc.setAttribute("w:val", "left");// 左对齐
		}

		var eleH = pPrs[0];
		eleH.setAttribute("w:val", "a");

		if (obj.tagName.indexOf('H') == 0) { //head
			eleH.setAttribute("w:val", obj.tagName.replace('H', ''));
		}
		var f = this.createElesLvl("w:rFonts", pPrs[2])[0];
		f.setAttribute("w:hint", "eastAsia");
		this.lastWR = null;

		this.setAtts(elep, {
			"w:rsidR" : "003D2D54",
			"w14:textId" : "77777777",
			"w14:paraId" : this.getParaId(),
			"w:rsidRDefault" : "0057281A"
		});
		return elep;
	};
	/**
	 *  <w:r>
	            <w:rPr>
	                <w:rFonts w:hint="eastAsia"/>
	                <w:noProof/>
	            </w:rPr>
	            <w:drawing>
	                <wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="271B1DC6" wp14:editId="7E056E7F">
	                    <wp:extent cx="358820" cy="360000"/>
	                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
	                    <wp:docPr id="1" name="图片 1"/>
	                    <wp:cNvGraphicFramePr>
	                        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	                                             noChangeAspect="1"/>
	                    </wp:cNvGraphicFramePr>
	                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
	                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
	                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
	                                <pic:nvPicPr>
	                                    <pic:cNvPr id="0" name="1.gif"/>
	                                    <pic:cNvPicPr/>
	                                </pic:nvPicPr>
	                                <pic:blipFill rotWithShape="1">
	                                    <a:blip r:embed="rId8">
	                                        <a:extLst>
	                                            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
	                                                <a14:useLocalDpi
	                                                        xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
	                                                        val="0"/>
	                                            </a:ext>
	                                        </a:extLst>
	                                    </a:blip>
	                                    <a:srcRect l="-301" t="-466" r="-301" b="-466"/>
	                                    <a:stretch/>
	                                </pic:blipFill>
	                                <pic:spPr bwMode="auto">
	                                    <a:xfrm>
	                                        <a:off x="0" y="0"/>
	                                        <a:ext cx="362156" cy="363346"/>
	                                    </a:xfrm>
	                                    <a:prstGeom prst="rect">
	                                        <a:avLst/>
	                                    </a:prstGeom>
	                                    <a:ln>
	                                        <a:noFill/>
	                                    </a:ln>
	                                    <a:extLst>
	                                        <a:ext uri="{53640926-AAD7-44d8-BBD7-CCE9431645EC}">
	                                            <a14:shadowObscured
	                                                    xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"/>
	                                        </a:ext>
	                                    </a:extLst>
	                                </pic:spPr>
	                            </pic:pic>
	                        </a:graphicData>
	                    </a:graphic>
	                </wp:inline>
	            </w:drawing>
	        </w:r>
	 * @param {Object} obj
	 * @memberOf {TypeName} 
	 */
	this.createPic = function(obj) {
		var o = $(obj);
		var picName = obj.src;
		this.pics.push(obj.src);

		var idx = picName.lastIndexOf('/');
		picName = picName.substring(idx + 1);
		var emuX = this.px2emu(o.width());
		var emuY = this.px2emu(o.height());

		var rs = this.createElesLvl("w:r,w:rPr,w:noProof");
		var wr = rs[0];
		var wdraws = this.createElesLvl("w:drawing,wp:inline,wp:extent");
		wr.appendChild(wdraws[0]);
		wExtent = wdraws[2];
		this.setAtts(wExtent, {
			'cx' : emuX,
			'cy' : emuY
		});

		var wpInline = wdraws[1];
		this.setAtts(wpInline, {
			distT : "0",
			distB : "0",
			distL : "0",
			distR : "0"
		});
		var eles = this.createElesSameLvl(
				'wp:effectExtent,wp:docPr,wp:cNvGraphicFramePr', wpInline);
		this.setAtts(eles[0], {
			l : "0",
			t : "0",
			r : "0",
			b : "0"
		});
		this.setAtts(eles[1], {
			id : "0",
			name : picName
		});
		var a_graphicFrameLocks = this.createEleNs('a:graphicFrameLocks',
				'http://schemas.openxmlformats.org/drawingml/2006/main');
		eles[2].appendChild(a_graphicFrameLocks);
		this.setAtts(a_graphicFrameLocks, {
			noChangeAspect : 1
		});
		//<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
		//	<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
		//		<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">

		var a_graphic = this.createEleNs('a:graphic',
				"http://schemas.openxmlformats.org/drawingml/2006/main");
		var a_graphicData = this.createEle('a:graphicData');
		a_graphicData.setAttribute('uri',
				"http://schemas.openxmlformats.org/drawingml/2006/picture");
		var pic_pic = this.createEleNs('pic:pic',
				"http://schemas.openxmlformats.org/drawingml/2006/picture");
		a_graphic.appendChild(a_graphicData);
		a_graphicData.appendChild(pic_pic);
		wpInline.appendChild(a_graphic);

		var pics = this.createElesSameLvl('pic:nvPicPr,pic:blipFill,pic:spPr',
				pic_pic);
		this.setAtts(pics[1], {
			'rotWithShape' : "1"
		});
		this.setAtts(pics[2], {
			'bwMode' : "auto"
		});
		//<pic:nvPicPr> 
		//	<pic:cNvPr id="0" name="1.gif"/>
		//	<pic:cNvPicPr/>
		var nvPicPrs = this
				.createElesSameLvl('pic:cNvPr,pic:cNvPicPr', pics[0]);
		this.setAtts(nvPicPrs[0], {
			id : "0",
			name : picName
		});
		//<a:blip r:embed="rId8">
		//        <a:extLst>
		//            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
		//                 <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
		//             </a:ext>
		//        </a:extLst>
		//</a:blip>
		//<a:srcRect l="-301" t="-466" r="-301" b="-466"/>
		//<a:stretch/>
		var pic_blipFills = this.createElesSameLvl(
				'a:blip,a:srcRect,a:stretch', pics[1]);
		this.setAtts(pic_blipFills[0], {
			'r:embed' : "pic" + (this.pics.length - 1)
		});
		this.setAtts(pic_blipFills[1], {
			l : "0",
			t : "0",
			r : "0",
			b : "0"
		});
		var a_extLsts = this.createEles('a:extLst,a:ext');
		pic_blipFills[0].appendChild(a_extLsts[0]);
		a_extLsts[1].setAttribute('uri',
				"{28A0092B-C50C-407E-A947-70E740481C1C}");

		var a14_useLocalDpi = this.createEleNs('a14:useLocalDpi',
				'http://schemas.microsoft.com/office/drawing/2010/main');
		a14_useLocalDpi.setAttribute('val', 0);
		a_extLsts[1].appendChild(a14_useLocalDpi);
		//		<a:xfrm>
		//		        <a:off x="0" y="0"/>
		//		        <a:ext cx="362156" cy="363346"/>
		//		</a:xfrm>
		//		<a:prstGeom prst="rect">
		//		        <a:avLst/>
		//		</a:prstGeom>
		//		<a:ln>
		//		       <a:noFill/>
		//		</a:ln>
		//		<a:extLst>
		//		       <a:ext uri="{53640926-AAD7-44d8-BBD7-CCE9431645EC}">
		//		              <a14:shadowObscured xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"/>
		//		        </a:ext>
		//		</a:extLst>

		var pic_spPrs = this.createElesSameLvl(
				'a:xfrm,a:prstGeom,a:ln,a:extLst', pics[2]);
		this.setAtts(pic_spPrs[1], {
			prst : "rect"
		});
		var axfrms = this.createElesSameLvl('a:off,a:ext', pic_spPrs[0]);
		this.setAtts(axfrms[0], {
			x : 0,
			y : 0
		});
		this.setAtts(axfrms[1], {
			cx : emuX,
			cy : emuY
		});
		this.createElesSameLvl('a:avLst', pic_spPrs[1]);
		this.createElesSameLvl('a:noFill', pic_spPrs[2]);
		//		var a_ext = this.createElesSameLvl('a:ext', pic_spPrs[3])[0];
		//		a_ext.setAttribute('uri', '{53640926-AAD7-44d8-BBD7-CCE9431645EC}');
		//		var a14_shadowObscured = this.createEleNs("a14:shadowObscured",
		//				"http://schemas.microsoft.com/office/drawing/2010/main");
		//		a_ext.appendChild(a14_shadowObscured);

		return wr;
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
			this.docks.push(this.doc.XmlDoc.getElementsByTagName('w:body')[0]);
		} else {
			this.docks.push(this.doc.XmlDoc.childNodes[0].childNodes[0]);
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
				//console.log(this.docks);
				break;
			}
		}
	};
	this.getDockTable = function() {
		for ( var i = this.docks.length - 1; i >= 0; i--) {
			var o = this.docks[i];
			if (o.tagName == 'w:body' || o.tagName == 'w:tc') {
				return o;
			}
		}
	};
	/**
	 * 包括英制单位（914,400 个 EMU 单位为 1 英寸）
	 * @param {Object} v
	 * @return {TypeName} 
	 */
	this.px2emu = function(v) {
		return parseInt(v / 96 * 914400);
	}
	this.EWA_DocTmp = {
		document : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
				<w:document\
				xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"\
				xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main"\
				xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"\
				xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office"\
				xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"\
				xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"\
				xmlns:v="urn:schemas-microsoft-com:vml"\
				xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"\
				xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"\
				xmlns:w10="urn:schemas-microsoft-com:office:word"\
				xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"\
				xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"\
				xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"\
				xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"\
				xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"\
				xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"\
				mc:Ignorable="w14 wp14"><w:body></w:body></w:document>',
		r : '<w:r w:rsidR="00A0061C" w:rsidRPr="00A0061C"><w:rPr>\
				<w:rFonts w:ascii="[FONT]" w:eastAsia="[FONT]" w:hAnsi="[FONT]" w:hint="eastAsia" />\
				<w:b />\
				<w:i />\
				<w:color w:val="FF0000" />\
				<w:sz w:val="36" />\
				<w:szCs w:val="36" />\
				<w:u w:val="single" />\
				</w:rPr>\
				<w:t>看看</w:t>\
				</w:r>',
		f : {
			"9px" : "7pt",
			"10px" : "7.5pt",
			"11px" : "8.5pt",
			"12px" : "9pt",
			"13px" : "10pt",
			"14px" : "10.5pt",
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
//var word = new EWA_DocWordClass();
//word.walker(oo);
//word.xml();
