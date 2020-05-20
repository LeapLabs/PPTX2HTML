'use strict'
import JSZip from 'jszip'
import tXml from './txml'

function base64ArrayBuffer(arrayBuff) {
    const buff = new Uint8Array(arrayBuff)
    let text = ''
    for (let i = 0; i < buff.byteLength; i++) {
        text += String.fromCharCode(buff[i])
    }
    return btoa(text)
}

function extractFileExtension(filename) {
    const dot = filename.lastIndexOf('.')
    if (dot === 0 || dot === -1) return ''
    return filename.substr(filename.lastIndexOf('.') + 1)
}

export default function processPptx(setOnMessage = () => {
}, postMessage) {
    const charts = []
    let chartID = 0

    let themeContent = null

    let slideLayoutClrOvride = ''

    const styleTable = {}

    let tableStyles

    setOnMessage(async e => {
        switch (e.type) {
            case 'processPPTX': {
                try {
                    await processPPTX(e.data)
                } catch (e) {
                    console.error('AN ERROR HAPPENED DURING processPPTX', e)
                    postMessage({
                        type: 'ERROR',
                        data: e.toString()
                    })
                }
                break
            }
            default:
        }
    })

    async function processPPTX(data) {
        const zip = JSZip(data)
        const dateBefore = new Date()

        const filesInfo = await getContentTypes(zip)
        const slideSize = await getSlideSize(zip)
        themeContent = await loadTheme(zip)

        tableStyles = await readXmlFile(zip, 'ppt/tableStyles.xml')

        postMessage({
            'type': 'slideSize',
            'data': slideSize
        })

        const numOfSlides = filesInfo['slides'].length
        for (let i = 0; i < numOfSlides; i++) {
            const filename = filesInfo['slides'][i]
            const slideHtml = await processSingleSlide(zip, filename, i, slideSize)
            postMessage({
                'type': 'slide',
                'data': slideHtml
            })
            postMessage({
                'type': 'progress-update',
                'data': (i + 1) * 100 / numOfSlides
            })
        }

        postMessage({
            'type': 'globalCSS',
            'data': genGlobalCSS()
        })

        const dateAfter = new Date()
        postMessage({
            'type': 'Done',
            'data': {
                time: dateAfter - dateBefore,
                slideSize: slideSize,
                charts
            }
        })
    }

    async function readXmlFile(zip, filename) {
        let xmlData = tXml(zip.file(filename).asText(), {simplify: 1})
        if (xmlData['?xml'] !== undefined) {
            return xmlData['?xml']
        } else {
            return xmlData
        }
    }

    async function getContentTypes(zip) {
        const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
        // console.log('CONTENT TYPES JSON', ContentTypesJson)
        const subObj = ContentTypesJson['Types']['Override']
        const slidesLocArray = []
        const slideLayoutsLocArray = []
        for (let i = 0; i < subObj.length; i++) {
            switch (subObj[i]['attrs']['ContentType']) {
                case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
                    slidesLocArray.push(subObj[i]['attrs']['PartName'].substr(1))
                    break
                case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
                    slideLayoutsLocArray.push(subObj[i]['attrs']['PartName'].substr(1))
                    break
                default:
            }
        }
        return {
            'slides': slidesLocArray,
            'slideLayouts': slideLayoutsLocArray
        }
    }

    async function getSlideSize(zip) {
        // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
        const content = await readXmlFile(zip, 'ppt/presentation.xml')
        const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
        return {
            'width': parseInt(sldSzAttrs['cx']) * 96 / 914400,
            'height': parseInt(sldSzAttrs['cy']) * 96 / 914400
        }
    }

    async function loadTheme(zip) {
        const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
        const relationshipArray = preResContent['Relationships']['Relationship']
        let themeURI
        if (relationshipArray.constructor === Array) {
            for (let i = 0; i < relationshipArray.length; i++) {
                if (relationshipArray[i]['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
                    themeURI = relationshipArray[i]['attrs']['Target']
                    break
                }
            }
        } else if (relationshipArray['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
            themeURI = relationshipArray['attrs']['Target']
        }

        if (themeURI === undefined) {
            throw Error('Can\'t open theme file.')
        }

        return readXmlFile(zip, 'ppt/' + themeURI)
    }

    async function processSingleSlide(zip, sldFileName, index, slideSize) {
        postMessage({
            'type': 'INFO',
            'data': 'Processing slide' + (index + 1)
        })

        // =====< Step 1 >=====
        // Read relationship filename of the slide (Get slideLayoutXX.xml)
        // @sldFileName: ppt/slides/slide1.xml
        // @resName: ppt/slides/_rels/slide1.xml.rels
        const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
        const resContent = await readXmlFile(zip, resName)
        let RelationshipArray = resContent['Relationships']['Relationship']
        let layoutFilename = ''
        const slideResObj = {}
        if (RelationshipArray.constructor === Array) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]['attrs']['Type']) {
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
                        layoutFilename = RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        break
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
                    default: {
                        slideResObj[RelationshipArray[i]['attrs']['Id']] = {
                            'type': RelationshipArray[i]['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
                            'target': RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        }
                    }
                }
            }
        } else {
            layoutFilename = RelationshipArray['attrs']['Target'].replace('../', 'ppt/')
        }
        // console.log(slideResObj);
        // Open slideLayoutXX.xml
        const slideLayoutContent = await readXmlFile(zip, layoutFilename)
        const slideLayoutTables = indexNodes(slideLayoutContent)
        const sldLayoutClrOvr = slideLayoutContent['p:sldLayout']['p:clrMapOvr']['a:overrideClrMapping']

        // console.log(slideLayoutClrOvride);
        if (sldLayoutClrOvr !== undefined) {
            slideLayoutClrOvride = sldLayoutClrOvr['attrs']
        }
        // =====< Step 2 >=====
        // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
        // @resName: ppt/slideLayouts/slideLayout1.xml
        // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
        const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
        const slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename)
        RelationshipArray = slideLayoutResContent['Relationships']['Relationship']
        let masterFilename = ''
        const layoutResObj = {}
        if (RelationshipArray.constructor === Array) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]['attrs']['Type']) {
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
                        masterFilename = RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        break
                    default:
                        layoutResObj[RelationshipArray[i]['attrs']['Id']] = {
                            'type': RelationshipArray[i]['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
                            'target': RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        }
                }
            }
        } else {
            masterFilename = RelationshipArray['attrs']['Target'].replace('../', 'ppt/')
        }
        // Open slideMasterXX.xml
        const slideMasterContent = await readXmlFile(zip, masterFilename)
        const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
        const slideMasterTables = indexNodes(slideMasterContent)

        // ///////////////Amir/////////////
        // Open slideMasterXX.xml.rels
        const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
        const slideMasterResContent = await readXmlFile(zip, slideMasterResFilename)
        RelationshipArray = slideMasterResContent['Relationships']['Relationship']
        let themeFilename = ''
        const masterResObj = {}
        if (RelationshipArray.constructor === Array) {
            for (let i = 0; i < RelationshipArray.length; i++) {
                switch (RelationshipArray[i]['attrs']['Type']) {
                    case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
                        themeFilename = RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        break
                    default:
                        masterResObj[RelationshipArray[i]['attrs']['Id']] = {
                            'type': RelationshipArray[i]['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
                            'target': RelationshipArray[i]['attrs']['Target'].replace('../', 'ppt/')
                        }
                }
            }
        } else {
            themeFilename = RelationshipArray['attrs']['Target'].replace('../', 'ppt/')
        }
        // console.log(themeFilename)
        // Load Theme file
        if (themeFilename !== undefined) {
            themeContent = await readXmlFile(zip, themeFilename)
        }
        // =====< Step 3 >=====
        const slideContent = await readXmlFile(zip, sldFileName)
        const nodes = slideContent['p:sld']['p:cSld']['p:spTree']
        const warpObj = {
            'zip': zip,
            'slideLayoutTables': slideLayoutTables,
            'slideMasterTables': slideMasterTables,
            'slideResObj': slideResObj,
            'slideMasterTextStyles': slideMasterTextStyles,
            'layoutResObj': layoutResObj,
            'masterResObj': masterResObj
        }

        let result = '<section>'

        for (let nodeKey in nodes) {
            if (nodes[nodeKey].constructor === Array) {
                for (let i = 0; i < nodes[nodeKey].length; i++) {
                    result += await processNodesInSlide(nodeKey, nodes[nodeKey][i], warpObj)
                }
            } else {
                result += await processNodesInSlide(nodeKey, nodes[nodeKey], warpObj)
            }
        }

        return result + '</section>'
    }

    function indexNodes(content) {
        const keys = Object.keys(content)
        const spTreeNode = content[keys[0]]['p:cSld']['p:spTree']

        const idTable = {}
        const idxTable = {}
        const typeTable = {}

        for (let key in spTreeNode) {
            if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') {
                continue
            }

            const targetNode = spTreeNode[key]

            if (targetNode.constructor === Array) {
                for (let i = 0; i < targetNode.length; i++) {
                    const nvSpPrNode = targetNode[i]['p:nvSpPr']
                    const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
                    const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
                    const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

                    if (id !== undefined) {
                        idTable[id] = targetNode[i]
                    }
                    if (idx !== undefined) {
                        idxTable[idx] = targetNode[i]
                    }
                    if (type !== undefined) {
                        typeTable[type] = targetNode[i]
                    }
                }
            } else {
                const nvSpPrNode = targetNode['p:nvSpPr']
                const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
                const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
                const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

                if (id !== undefined) {
                    idTable[id] = targetNode
                }
                if (idx !== undefined) {
                    idxTable[idx] = targetNode
                }
                if (type !== undefined) {
                    typeTable[type] = targetNode
                }
            }
        }

        return {'idTable': idTable, 'idxTable': idxTable, 'typeTable': typeTable}
    }

    async function processNodesInSlide(nodeKey, nodeValue, warpObj) {
        let result = ''

        switch (nodeKey) {
            case 'p:sp':    // Shape, Text
                result = processSpNode(nodeValue, warpObj)
                break
            case 'p:cxnSp':    // Shape, Text (with connection)
                result = processCxnSpNode(nodeValue, warpObj)
                break
            case 'p:pic':    // Picture
                result = processPicNode(nodeValue, warpObj)
                break
            case 'p:grpSp':    // 群組
                result = await processGroupSpNode(nodeValue, warpObj)
                break
            default:
        }

        return result
    }

    async function processGroupSpNode(node, warpObj) {
        const factor = 96 / 914400

        const xfrmNode = node['p:grpSpPr']['a:xfrm']
        const x = parseInt(xfrmNode['a:off']['attrs']['x']) * factor
        const y = parseInt(xfrmNode['a:off']['attrs']['y']) * factor
        const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * factor
        const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * factor
        const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * factor
        const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * factor
        const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * factor
        const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * factor

        const order = node['attrs']['order']

        let result = '<div class=\'block group\' style=\'z-index: ' + order + '; top: ' + (y - chy) + 'px; left: ' + (x - chx) + 'px; width: ' + (cx - chcx) + 'px; height: ' + (cy - chcy) + 'px;\'>'

        // Procsee all child nodes
        for (let nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (let i = 0; i < node[nodeKey].length; i++) {
                    result += await processNodesInSlide(nodeKey, node[nodeKey][i], warpObj)
                }
            } else {
                result += await processNodesInSlide(nodeKey, node[nodeKey], warpObj)
            }
        }

        result += '</div>'

        return result
    }

    function processSpNode(node, warpObj) {
        /*
    *  958    <xsd:complexType name="CT_GvmlShape">
    *  959   <xsd:sequence>
    *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
    *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
    *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
    *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
    *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    *  965   </xsd:sequence>
    *  966 </xsd:complexType>
    */

        const id = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id'])
        const name = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name'])
        const idx = (getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) === undefined) ? undefined : getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
        let type = (getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph']) === undefined) ? undefined : getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
        const order = getTextByPathList(node, ['attrs', 'order'])

        let slideLayoutSpNode = undefined
        let slideMasterSpNode = undefined

        if (type !== undefined) {
            if (idx !== undefined) {
                slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
                slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
            } else {
                slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
                slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
            }
        } else {
            if (idx !== undefined) {
                slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
                slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
            } else {
                // Nothing
            }
        }

        if (type === undefined) {
            type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
            if (type === undefined) {
                type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
            }
        }

        return genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj)
    }

    function processCxnSpNode(node, warpObj) {
        const id = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['id']
        const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']
        // const idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
        // const type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
        // <p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
        const order = node['attrs']['order']

        return genShape(node, undefined, undefined, id, name, undefined, undefined, order, warpObj)
    }

    function genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj) {
        const xfrmList = ['p:spPr', 'a:xfrm']
        const slideXfrmNode = getTextByPathList(node, xfrmList)
        const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
        const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

        let result = ''
        const shpId = getTextByPathList(node, ['attrs', 'order'])
        // console.log("shpId: ",shpId)
        const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])

        // custGeom - Amir
        const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

        let isFlipV = false
        if (getTextByPathList(slideXfrmNode, ['attrs', 'flipV']) === '1' || getTextByPathList(slideXfrmNode, ['attrs', 'flipH']) === '1') {
            isFlipV = true
        }
        // ///////////////////////Amir////////////////////////
        // rotate
        const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

        if (shapType !== undefined && custShapType === undefined) {
            result += '<div class=\'block content ' +
                '\' _id=\'' + id + '\' _idx=\'' + idx + '\' _type=\'' + type + '\' Name=\'' + name +
                '\' style=\'' +
                getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                ' z-index: ' + order + ';' +
                'transform: rotate(' + rotate + 'deg);' +
                '\'>'

            // TextBody
            if (node['p:txBody'] !== undefined) {
                result += genTextBody(node['p:txBody'], slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            }
            result += '</div>'
        } else if (custShapType !== undefined) {
            // custGeom here - Amir ///////////////////////////////////////////////////////
            // http://officeopenxml.com/drwSp-custGeom.php
            const pathLstNode = getTextByPathList(custShapType, ['a:pathLst'])
            // const pathNode = getTextByPathList(pathLstNode, ['a:path', 'attrs'])
            // const maxX = parseInt(pathNode['w']) * 96 / 914400
            // const maxY = parseInt(pathNode['h']) * 96 / 914400
            // console.log("w = "+w+"\nh = "+h+"\nmaxX = "+maxX +"\nmaxY = " + maxY);
            // cheke if it is close shape
            const closeNode = getTextByPathList(pathLstNode, ['a:path', 'a:close'])
            const startPoint = getTextByPathList(pathLstNode, ['a:path', 'a:moveTo', 'a:pt', 'attrs'])
            const spX = startPoint ? parseInt(startPoint['x']) * 96 / 914400 : 0;
            const spY = startPoint ? parseInt(startPoint['y']) * 96 / 914400 : 0;
            let d = 'M' + spX + ',' + spY
            const pathNodes = getTextByPathList(pathLstNode, ['a:path'])
            const lnToNodes = pathNodes['a:lnTo']
            const cubicBezToNodes = pathNodes['a:cubicBezTo']
            const sortblAry = []
            if (lnToNodes !== undefined) {
                Object.keys(lnToNodes).forEach(function (key) {
                    const lnToPtNode = lnToNodes[key]['a:pt']
                    if (lnToPtNode !== undefined) {
                        Object.keys(lnToPtNode).forEach(function (key2) {
                            const ptObj = {}
                            const lnToNoPt = lnToPtNode[key2]
                            const ptX = lnToNoPt['x']
                            const ptY = lnToNoPt['y']
                            const ptOrdr = lnToNoPt['order']
                            ptObj.type = 'lnto'
                            ptObj.order = ptOrdr
                            ptObj.x = ptX
                            ptObj.y = ptY
                            sortblAry.push(ptObj)
                            // console.log(key2, lnToNoPt);
                        })
                    }
                })
            }
            if (cubicBezToNodes !== undefined) {
                Object.keys(cubicBezToNodes).forEach(function (key) {
                    // console.log("cubicBezTo["+key+"]:");
                    const cubicBezToPtNodes = cubicBezToNodes[key]['a:pt']
                    if (cubicBezToPtNodes !== undefined) {
                        Object.keys(cubicBezToPtNodes).forEach(function (key2) {
                            // console.log("cubicBezTo["+key+"]pt["+key2+"]:");
                            const cubBzPts = cubicBezToPtNodes[key2]
                            Object.keys(cubBzPts).forEach(function (key3) {
                                // console.log(key3, cubBzPts[key3]);
                                const ptObj = {}
                                const cubBzPt = cubBzPts[key3]
                                const ptX = cubBzPt['x']
                                const ptY = cubBzPt['y']
                                const ptOrdr = cubBzPt['order']
                                ptObj.type = 'cubicBezTo'
                                ptObj.order = ptOrdr
                                ptObj.x = ptX
                                ptObj.y = ptY
                                sortblAry.push(ptObj)
                            })
                        })
                    }
                })
            }
            const sortByOrder = sortblAry.slice(0)
            sortByOrder.sort(function (a, b) {
                return a.order - b.order
            })
            // console.log(sortByOrder);
            let k = 0
            while (k < sortByOrder.length) {
                if (sortByOrder[k].type === 'lnto') {
                    const Lx = parseInt(sortByOrder[k].x) * 96 / 914400
                    const Ly = parseInt(sortByOrder[k].y) * 96 / 914400
                    d += 'L' + Lx + ',' + Ly
                    k++
                } else { // "cubicBezTo"
                    const Cx1 = parseInt(sortByOrder[k].x) * 96 / 914400
                    const Cy1 = parseInt(sortByOrder[k].y) * 96 / 914400
                    const Cx2 = parseInt(sortByOrder[k + 1].x) * 96 / 914400
                    const Cy2 = parseInt(sortByOrder[k + 1].y) * 96 / 914400
                    const Cx3 = parseInt(sortByOrder[k + 2].x) * 96 / 914400
                    const Cy3 = parseInt(sortByOrder[k + 2].y) * 96 / 914400

                    d += 'C' + Cx1 + ',' + Cy1 + ' ' + Cx2 + ',' + Cy2 + ' ' + Cx3 + ',' + Cy3
                    k += 3
                }
            }


            result += '<div class=\'block content\'>'

            // TextBody
            if (node['p:txBody'] !== undefined) {
                result += genTextBody(node['p:txBody'], slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            }
            result += '</div>'
        } else {
            result += '<div class=\'block content\'>'

            // TextBody
            if (node['p:txBody'] !== undefined) {
                result += genTextBody(node['p:txBody'], slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            }
            result += '</div>'
        }

        return result
    }

    function processPicNode(node, warpObj) {

        let rtrnData = ''
        let mediaPicFlag = false
        const order = node['attrs']['order']

        const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
        const imgName = warpObj['slideResObj'][rid]['target']
        const imgFileExt = extractFileExtension(imgName).toLowerCase()
        const zip = warpObj['zip']

        const imgArrayBuffer = zip.file(imgName).asArrayBuffer()
        let mimeType = ''
        const xfrmNode = node['p:spPr']['a:xfrm']
        // /////////////////////////////////////Amir//////////////////////////////
        // const rotate = angleToDegrees(node['p:spPr']['a:xfrm']['attrs']['rot'])

        let rotate = 0
        let rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
        if (rotateNode !== undefined) {
            rotate = angleToDegrees(rotateNode)
        }
        //video
        let vdoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
        let vdoRid, vdoFile, vdoFileExt, vdoMimeType, uInt8Array, blob, vdoBlob, mediaSupportFlag = false
        let mediaProcess = true
        if (vdoNode !== undefined & mediaProcess) {
            vdoRid = vdoNode['attrs']['r:link']
            vdoFile = warpObj['slideResObj'][vdoRid]['target']
            uInt8Array = zip.file(vdoFile).asArrayBuffer()
            vdoFileExt = extractFileExtension(vdoFile).toLowerCase()
            if (vdoFileExt === 'mp4' || vdoFileExt === 'webm' || vdoFileExt === 'ogg') {
                vdoMimeType = getMimeType(vdoFileExt)
                blob = new Blob([uInt8Array], {
                    type: vdoMimeType
                })
                vdoBlob = URL.createObjectURL(blob)
                mediaSupportFlag = true
                mediaPicFlag = true
            }
        }
        //Audio
        let audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
        let audioRid, audioFile, audioFileExt, audioMimeType, uInt8ArrayAudio, blobAudio, audioBlob
        let audioPlayerFlag = false
        let audioObjc
        if (audioNode !== undefined & mediaProcess) {
            audioRid = audioNode['attrs']['r:link']
            audioFile = warpObj['slideResObj'][audioRid]['target']
            audioFileExt = extractFileExtension(audioFile).toLowerCase()
            if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
                uInt8ArrayAudio = zip.file(audioFile).asArrayBuffer()
                blobAudio = new Blob([uInt8ArrayAudio])
                audioBlob = URL.createObjectURL(blobAudio)
                let cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * 20
                let cy = xfrmNode['a:ext']['attrs']['cy']
                let x = parseInt(xfrmNode['a:off']['attrs']['x']) / 2.5
                let y = xfrmNode['a:off']['attrs']['y']
                audioObjc = {
                    'a:ext': {
                        'attrs': {
                            'cx': cx,
                            'cy': cy
                        }
                    },
                    'a:off': {
                        'attrs': {
                            'x': x,
                            'y': y

                        }
                    }
                }
                audioPlayerFlag = true
                mediaSupportFlag = true
                mediaPicFlag = true
            }
        }
        // ////////////////////////////////////////////////////////////////////////
        // mimeType = getImageMimeType(imgFileExt)
        // return '<div class=\'block content\' style=\'' + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
        //   ' z-index: ' + order + ';' +
        //   'transform: rotate(' + rotate + 'deg);' +
        //   '\'><img src=\'data:' + mimeType + ';base64,' + base64ArrayBuffer(imgArrayBuffer) + '\' style=\'width: 100%; height: 100%\'/></div>'

        mimeType = getMimeType(imgFileExt)
        rtrnData = '<div class=\'block content\' style=\'' +
            ((mediaProcess && audioPlayerFlag) ? getPosition(audioObjc, undefined, undefined) : getPosition(xfrmNode, undefined, undefined)) +
            ((mediaProcess && audioPlayerFlag) ? getSize(audioObjc, undefined, undefined) : getSize(xfrmNode, undefined, undefined)) +
            ' z-index: ' + order + ';' +
            'transform: rotate(' + rotate + 'deg);\'>'
        if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
            rtrnData += '<img src=\'data:' + mimeType + ';base64,' + base64ArrayBuffer(imgArrayBuffer) + '\' style=\'width: 100%; height: 100%\'/>'
        } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
            if (vdoNode !== undefined) {
                rtrnData += '<video  src=\'' + vdoBlob + '\' controls style=\'width: 100%; height: 100%\'>Your browser does not support the video tag.</video>'
            }
            if (audioNode !== undefined) {
                rtrnData += '<audio id="audio_player" controls ><source src="' + audioBlob + '"></audio>'
                //'<button onclick="audio_player.play()">Play</button>'+
                //'<button onclick="audio_player.pause()">Pause</button>';
            }
        }
        if (!mediaSupportFlag && mediaPicFlag) {
            rtrnData += '<span style=\'color:red;font-size:40px;position: absolute;\'>This media file Not supported by HTML5</span>'
        }
        if ((vdoNode !== undefined || audioNode !== undefined) && !mediaProcess && mediaSupportFlag) {
            console.log('Founded supported media file but media process disabled (mediaProcess=false)')
        }
        rtrnData += '</div>'
        //console.log(rtrnData)
        return rtrnData

    }

    function genTextBody(textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
        let text = ''
        const slideMasterTextStyles = warpObj['slideMasterTextStyles']

        if (textBodyNode === undefined) {
            return text
        }
        // rtl : <p:txBody>
        //          <a:bodyPr wrap="square" rtlCol="1">

        // const rtlStr = "";
        let pNode
        let rNode
        if (textBodyNode['a:p'].constructor === Array) {
            // multi p
            for (let i = 0; i < textBodyNode['a:p'].length; i++) {
                pNode = textBodyNode['a:p'][i]
                rNode = pNode['a:r']

                // const isRTL = getTextDirection(pNode, type, slideMasterTextStyles);
                // rtlStr = "";//"dir='"+isRTL+"'";

                text += '<div class=\'' + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + '\'>'
                text += genBuChar(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)

                if (rNode === undefined) {
                    // without r
                    text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
                } else if (rNode.constructor === Array) {
                    // with multi r
                    for (let j = 0; j < rNode.length; j++) {
                        text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type, warpObj)
                        // ////////////////Amir////////////
                        if (pNode['a:br'] !== undefined) {
                            text += '<br>'
                        }
                        // ////////////////////////////////
                    }
                } else {
                    // with one r
                    text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
                }
                text += '</div>'
            }
        } else {
            // one p
            pNode = textBodyNode['a:p']
            rNode = pNode['a:r']

            // const isRTL = getTextDirection(pNode, type, slideMasterTextStyles);
            // rtlStr = "";//"dir='"+isRTL+"'";

            text += '<div class=\'slide-prgrph ' + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + '\'>'
            text += genBuChar(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            if (rNode === undefined) {
                // without r
                text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            } else if (rNode.constructor === Array) {
                // with multi r
                for (let j = 0; j < rNode.length; j++) {
                    text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type, warpObj)
                    // ////////////////Amir////////////
                    if (pNode['a:br'] !== undefined) {
                        text += '<br>'
                    }
                    // ////////////////////////////////
                }
            } else {
                // with one r
                text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj)
            }
            text += '</div>'
        }

        return text
    }

    function genBuChar(node, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
        // /////////////////////////////////////Amir///////////////////////////////
        const sldMstrTxtStyles = warpObj['slideMasterTextStyles']

        const rNode = node['a:r']
        let dfltBultColor, dfltBultSize, bultColor, bultSize
        if (rNode !== undefined) {
            dfltBultColor = getFontColor(rNode, type, sldMstrTxtStyles)
            dfltBultSize = getFontSize(rNode, slideLayoutSpNode, slideMasterSpNode, type, sldMstrTxtStyles)
        } else {
            dfltBultColor = getFontColor(node, type, sldMstrTxtStyles)
            dfltBultSize = getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, sldMstrTxtStyles)
        }
        // console.log("Bullet Size: " + bultSize);

        let bullet = ''
        // ///////////////////////////////////////////////////////////////

        const pPrNode = node['a:pPr']

        // ////////////////cheke if is rtl ///Amir ////////////////////////////////////
        const getRtlVal = getTextByPathList(pPrNode, ['attrs', 'rtl'])
        let isRTL = false
        if (getRtlVal !== undefined && getRtlVal === '1') {
            isRTL = true
        }
        // //////////////////////////////////////////////////////////

        let lvl = parseInt(getTextByPathList(pPrNode, ['attrs', 'lvl']))
        if (isNaN(lvl)) {
            lvl = 0
        }

        const buChar = getTextByPathList(pPrNode, ['a:buChar', 'attrs', 'char'])
        // ///////////////////////////////Amir///////////////////////////////////
        let buType = 'TYPE_NONE'
        const buNum = getTextByPathList(pPrNode, ['a:buAutoNum', 'attrs', 'type'])
        const buPic = getTextByPathList(pPrNode, ['a:buBlip'])
        if (buChar !== undefined) {
            buType = 'TYPE_BULLET'
            // console.log("Bullet Chr to code: " + buChar.charCodeAt(0));
        }
        if (buNum !== undefined) {
            buType = 'TYPE_NUMERIC'
        }
        if (buPic !== undefined) {
            buType = 'TYPE_BULPIC'
        }

        let buFontAttrs
        if (buType !== 'TYPE_NONE') {
            buFontAttrs = getTextByPathList(pPrNode, ['a:buFont', 'attrs'])
        }
        // console.log("Bullet Type: " + buType);
        // console.log("NumericTypr: " + buNum);
        // console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
        // get definde bullet COLOR
        let defBultColor = 'NoNe'

        if (pPrNode) {
            const buClrNode = pPrNode['a:buClr']
            if (buClrNode !== undefined) {
                defBultColor = getSolidFill(buClrNode)
            } else {
                // console.log("buClrNode: " + buClrNode);
            }
        }

        if (defBultColor === 'NoNe') {
            bultColor = dfltBultColor
        } else {
            bultColor = '#' + defBultColor
        }
        // get definde bullet SIZE
        let buFontSize
        buFontSize = getTextByPathList(pPrNode, ['a:buSzPts', 'attrs', 'val']) // pt
        if (buFontSize !== undefined) {
            bultSize = parseInt(buFontSize) / 100 + 'pt'
        } else {
            buFontSize = getTextByPathList(pPrNode, ['a:buSzPct', 'attrs', 'val'])
            if (buFontSize !== undefined) {
                const prcnt = parseInt(buFontSize) / 100000
                // dfltBultSize = XXpt
                const dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2)
                bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + 'pt'
            } else {
                bultSize = dfltBultSize
            }
        }
        // //////////////////////////////////////////////////////////////////////
        let marginLeft
        let marginRight
        if (buType === 'TYPE_BULLET') {
            // const buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
            if (buFontAttrs !== undefined) {
                marginLeft = parseInt(getTextByPathList(pPrNode, ['attrs', 'marL'])) * 96 / 914400
                marginRight = parseInt(buFontAttrs['pitchFamily'])
                if (isNaN(marginLeft)) {
                    marginLeft = 328600 * 96 / 914400
                }
                if (isNaN(marginRight)) {
                    marginRight = 0
                }
                const typeface = buFontAttrs['typeface']

                bullet = '<span style=\'font-family: ' + typeface +
                    '; margin-left: ' + marginLeft * lvl + 'px' +
                    '; margin-right: ' + marginRight + 'px' +
                    ';color:' + bultColor +
                    ';font-size:' + bultSize + ';'
                if (isRTL) {
                    bullet += ' float: right;  direction:rtl'
                }
                bullet += '\'>' + buChar + '</span>'
            } else {
                marginLeft = 328600 * 96 / 914400 * lvl

                bullet = '<span style=\'margin-left: ' + marginLeft + 'px;\'>' + buChar + '</span>'
            }
        } else if (buType === 'TYPE_NUMERIC') { // /////////Amir///////////////////////////////
            if (buFontAttrs !== undefined) {
                marginLeft = parseInt(getTextByPathList(pPrNode, ['attrs', 'marL'])) * 96 / 914400
                marginRight = parseInt(buFontAttrs['pitchFamily'])

                if (isNaN(marginLeft)) {
                    marginLeft = 328600 * 96 / 914400
                }
                if (isNaN(marginRight)) {
                    marginRight = 0
                }
                // const typeface = buFontAttrs["typeface"];

                bullet = '<span style=\'margin-left: ' + marginLeft * lvl + 'px' +
                    '; margin-right: ' + marginRight + 'px' +
                    ';color:' + bultColor +
                    ';font-size:' + bultSize + ';'
                if (isRTL) {
                    bullet += ' float: right; direction:rtl;'
                } else {
                    bullet += ' float: left; direction:ltr;'
                }
                bullet += '\' data-bulltname = \'' + buNum + '\' data-bulltlvl = \'' + lvl + '\' class=\'numeric-bullet-style\'></span>'
            } else {
                marginLeft = 328600 * 96 / 914400 * lvl
                bullet = '<span style=\'margin-left: ' + marginLeft + 'px;'
                if (isRTL) {
                    bullet += ' float: right; direction:rtl;'
                } else {
                    bullet += ' float: left; direction:ltr;'
                }
                bullet += '\' data-bulltname = \'' + buNum + '\' data-bulltlvl = \'' + lvl + '\' class=\'numeric-bullet-style\'></span>'
            }
        } else if (buType === 'TYPE_BULPIC') { // PIC BULLET
            marginLeft = parseInt(getTextByPathList(pPrNode, ['attrs', 'marL'])) * 96 / 914400
            marginRight = parseInt(getTextByPathList(pPrNode, ['attrs', 'marR'])) * 96 / 914400

            if (isNaN(marginRight)) {
                marginRight = 0
            }
            // console.log("marginRight: "+marginRight)
            // buPic
            if (isNaN(marginLeft)) {
                marginLeft = 328600 * 96 / 914400
            } else {
                marginLeft = 0
            }
            // const buPicId = getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
            const buPicId = getTextByPathList(buPic, ['a:blip', 'attrs', 'r:embed'])
            // const svgPicPath = ''
            let buImg
            if (buPicId !== undefined) {
                // svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                // buImg = warpObj["zip"].file(svgPicPath).asText();
                // }else{
                // buPicId = getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                const imgPath = warpObj['slideResObj'][buPicId]['target']
                const imgArrayBuffer = warpObj['zip'].file(imgPath).asArrayBuffer()
                const imgExt = imgPath.split('.').pop()
                const imgMimeType = getImageMimeType(imgExt)
                buImg = '<img src=\'data:' + imgMimeType + ';base64,' + base64ArrayBuffer(imgArrayBuffer) + '\' style=\'width: 100%; height: 100%\'/>'
                // console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
            }
            if (buPicId === undefined) {
                buImg = '&#8227;'
            }
            bullet = '<span style=\'margin-left: ' + marginLeft * lvl + 'px' +
                '; margin-right: ' + marginRight + 'px' +
                ';width:' + bultSize + ';display: inline-block; '
            if (isRTL) {
                bullet += ' float: right;direction:rtl'
            }
            bullet += '\'>' + buImg + '  </span>'
            // ////////////////////////////////////////////////////////////////////////////////////
        } else {
            bullet = '<span style=\'margin-left: ' + 328600 * 96 / 914400 * lvl + 'px' +
                '; margin-right: ' + 0 + 'px;\'></span>'
        }

        return bullet
    }

    function genSpanElement(node, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
        const slideMasterTextStyles = warpObj['slideMasterTextStyles']

        let text = node['a:t']
        if (typeof text !== 'string' && !(text instanceof String)) {
            text = getTextByPathList(node, ['a:fld', 'a:t'])
            if (typeof text !== 'string' && !(text instanceof String)) {
                text = '&nbsp;'
            }
        }

        let styleText =
            'color:' + getFontColor(node, type, slideMasterTextStyles) +
            ';font-size:' + getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) +
            ';font-family:' + getFontType(node, type, slideMasterTextStyles) +
            ';font-weight:' + getFontBold(node, type, slideMasterTextStyles) +
            ';font-style:' + getFontItalic(node, type, slideMasterTextStyles) +
            ';text-decoration:' + getFontDecoration(node, type, slideMasterTextStyles) +
            ';text-align:' + getTextHorizontalAlign(node, type, slideMasterTextStyles) +
            ';vertical-align:' + getTextVerticalAlign(node, type, slideMasterTextStyles) +
            ';'
        // ////////////////Amir///////////////
        const highlight = getTextByPathList(node, ['a:rPr', 'a:highlight'])
        if (highlight !== undefined) {
            styleText += 'background-color:#' + getSolidFill(highlight) + ';'
            styleText += 'Opacity:' + getColorOpacity(highlight) + ';'
        }
        // /////////////////////////////////////////
        let cssName = ''

        if (styleText in styleTable) {
            cssName = styleTable[styleText]['name']
        } else {
            cssName = '_css_' + (Object.keys(styleTable).length + 1)
            styleTable[styleText] = {
                'name': cssName,
                'text': styleText
            }
        }

        const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
        // get link colors : TODO
        if (linkID !== undefined) {
            const linkURL = warpObj['slideResObj'][linkID]['target']
            return '<span class=\'text-block ' + cssName + '\'><a href=\'' + linkURL + '\' target=\'_blank\'>' + text.replace(/\s/i, '&nbsp;') + '</a></span>'
        } else {
            return '<span class=\'text-block ' + cssName + '\'>' + text.replace(/\s/i, '&nbsp;') + '</span>'
        }
    }

    function genGlobalCSS() {
        let cssText = ''
        for (let key in styleTable) {
            cssText += 'section .' + styleTable[key]['name'] + '{' + styleTable[key]['text'] + '}\n'
        }
        return cssText
    }

    function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
        let off
        let x = -1
        let y = -1

        if (slideSpNode !== undefined) {
            off = slideSpNode['a:off']['attrs']
        } else if (slideLayoutSpNode !== undefined) {
            off = slideLayoutSpNode['a:off']['attrs']
        } else if (slideMasterSpNode !== undefined) {
            off = slideMasterSpNode['a:off']['attrs']
        }

        if (off === undefined) {
            return ''
        } else {
            x = parseInt(off['x']) * 96 / 914400
            y = parseInt(off['y']) * 96 / 914400
            return (isNaN(x) || isNaN(y)) ? '' : 'top:' + y + 'px; left:' + x + 'px;'
        }
    }

    function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
        let ext
        let w = -1
        let h = -1

        if (slideSpNode !== undefined) {
            ext = slideSpNode['a:ext']['attrs']
        } else if (slideLayoutSpNode !== undefined) {
            ext = slideLayoutSpNode['a:ext']['attrs']
        } else if (slideMasterSpNode !== undefined) {
            ext = slideMasterSpNode['a:ext']['attrs']
        }

        if (ext === undefined) {
            return ''
        } else {
            w = parseInt(ext['cx']) * 96 / 914400
            h = parseInt(ext['cy']) * 96 / 914400
            return (isNaN(w) || isNaN(h)) ? '' : 'width:' + w + 'px; height:' + h + 'px;'
        }
    }

    function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
        let algn = getTextByPathList(node, ['a:pPr', 'attrs', 'algn'])
        if (algn === undefined) {
            algn = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
            if (algn === undefined) {
                algn = getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
                if (algn === undefined) {
                    switch (type) {
                        case 'title':
                        case 'subTitle':
                        case 'ctrTitle': {
                            algn = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'attrs', 'alng'])
                            break
                        }
                        default: {
                            algn = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'attrs', 'alng'])
                        }
                    }
                }
            }
        }
        // TODO:
        if (algn === undefined) {
            if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
                return 'h-mid'
            } else if (type === 'sldNum') {
                return 'h-right'
            }
        }
        return algn === 'ctr' ? 'h-mid' : algn === 'r' ? 'h-right' : 'h-left'
    }

    function getFontType(node, type) {
        let typeface = getTextByPathList(node, ['a:rPr', 'a:latin', 'attrs', 'typeface'])

        if (typeface === undefined) {
            const fontSchemeNode = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:fontScheme'])
            if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
                typeface = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
            } else if (type === 'body') {
                typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
            } else {
                typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
            }
        }

        return (typeface === undefined) ? 'inherit' : typeface
    }

    function getFontColor(node, type, slideMasterTextStyles) {
        const solidFillNode = getTextByPathStr(node, 'a:rPr a:solidFill')

        const color = getSolidFill(solidFillNode)
        // console.log(themeContent)
        // const schemeClr = getTextByPathList(buClrNode ,["a:schemeClr", "attrs","val"]);
        return (color === undefined || color === 'FFF') ? '#000' : '#' + color
    }

    function getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
        let fontSize
        let sz
        if (node['a:rPr'] !== undefined) {
            fontSize = parseInt(node['a:rPr']['attrs']['sz']) / 100
        }

        if ((isNaN(fontSize) || fontSize === undefined)) {
            sz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
            fontSize = parseInt(sz) / 100
        }

        if (isNaN(fontSize) || fontSize === undefined) {
            if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
                sz = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
            } else if (type === 'body') {
                sz = getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
            } else if (type === 'dt' || type === 'sldNum') {
                sz = '1200'
            } else if (type === undefined) {
                sz = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
            }
            fontSize = parseInt(sz) / 100
        }

        const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
        if (baseline !== undefined && !isNaN(fontSize)) {
            fontSize -= 10
        }

        return isNaN(fontSize) ? 'inherit' : (fontSize + 'pt')
    }

    function getFontBold(node, type, slideMasterTextStyles) {
        return (node['a:rPr'] !== undefined && node['a:rPr']['attrs']['b'] === '1') ? 'bold' : 'initial'
    }

    function getFontItalic(node, type, slideMasterTextStyles) {
        return (node['a:rPr'] !== undefined && node['a:rPr']['attrs']['i'] === '1') ? 'italic' : 'normal'
    }

    function getFontDecoration(node, type, slideMasterTextStyles) {
        // /////////////////////////////Amir///////////////////////////////
        if (node['a:rPr'] !== undefined) {
            const underLine = node['a:rPr']['attrs']['u'] !== undefined ? node['a:rPr']['attrs']['u'] : 'none'
            const strikethrough = node['a:rPr']['attrs']['strike'] !== undefined ? node['a:rPr']['attrs']['strike'] : 'noStrike'
            // console.log("strikethrough: "+strikethrough);

            if (underLine !== 'none' && strikethrough === 'noStrike') {
                return 'underline'
            } else if (underLine === 'none' && strikethrough !== 'noStrike') {
                return 'line-through'
            } else if (underLine !== 'none' && strikethrough !== 'noStrike') {
                return 'underline line-through'
            } else {
                return 'initial'
            }
        } else {
            return 'initial'
        }
        // ///////////////////////////////////////////////////////////////
        // return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
    }

// //////////////////////////////////Amir/////////////////////////////////////
    function getTextHorizontalAlign(node, type, slideMasterTextStyles) {
        const getAlgn = getTextByPathList(node, ['a:pPr', 'attrs', 'algn'])
        let align = 'initial'
        if (getAlgn !== undefined) {
            switch (getAlgn) {
                case 'l': {
                    align = 'left'
                    break
                }
                case 'r': {
                    align = 'right'
                    break
                }
                case 'ctr': {
                    align = 'center'
                    break
                }
                case 'just': {
                    align = 'justify'
                    break
                }
                case 'dist': {
                    align = 'justify'
                    break
                }
                default:
                    align = 'initial'
            }
        }
        return align
    }

// ///////////////////////////////////////////////////////////////////
    function getTextVerticalAlign(node, type, slideMasterTextStyles) {
        const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
        return baseline === undefined ? 'baseline' : (parseInt(baseline) / 1000) + '%'
    }

    function getSolidFill(node) {
        if (node === undefined) {
            return undefined
        }

        let color = 'FFF'

        if (node['a:srgbClr'] !== undefined) {
            color = getTextByPathList(node, ['a:srgbClr', 'attrs', 'val']) // #...
        } else if (node['a:schemeClr'] !== undefined) { // a:schemeClr
            const schemeClr = getTextByPathList(node, ['a:schemeClr', 'attrs', 'val'])
            // console.log(schemeClr)
            color = getSchemeColorFromTheme('a:' + schemeClr, undefined) // #...
        } else if (node['a:scrgbClr'] !== undefined) {
            // <a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
            const defBultColorVals = node['a:scrgbClr']['attrs']
            const red = (defBultColorVals['r'].indexOf('%') !== -1) ? defBultColorVals['r'].split('%').shift() : defBultColorVals['r']
            const green = (defBultColorVals['g'].indexOf('%') !== -1) ? defBultColorVals['g'].split('%').shift() : defBultColorVals['g']
            const blue = (defBultColorVals['b'].indexOf('%') !== -1) ? defBultColorVals['b'].split('%').shift() : defBultColorVals['b']
            // const scrgbClr = red + ',' + green + ',' + blue
            color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100))
            // console.log("scrgbClr: " + scrgbClr);
        } else if (node['a:prstClr'] !== undefined) {
            // <a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
            const prstClr = node['a:prstClr']['attrs']['val']
            color = getColorName2Hex(prstClr)
            // console.log("prstClr: " + prstClr+" => hexClr: "+color);
        } else if (node['a:hslClr'] !== undefined) {
            // <a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
            const defBultColorVals = node['a:hslClr']['attrs']
            const hue = Number(defBultColorVals['hue']) / 100000
            const sat = Number((defBultColorVals['sat'].indexOf('%') !== -1) ? defBultColorVals['sat'].split('%').shift() : defBultColorVals['sat']) / 100
            const lum = Number((defBultColorVals['lum'].indexOf('%') !== -1) ? defBultColorVals['lum'].split('%').shift() : defBultColorVals['lum']) / 100
            // const hslClr = defBultColorVals['hue'] + ',' + defBultColorVals['sat'] + ',' + defBultColorVals['lum']
            const hsl2rgb = hslToRgb(hue, sat, lum)
            color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b)
            // defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
            // console.log("hslClr: " + hslClr);
        } else if (node['a:sysClr'] !== undefined) {
            // <a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
            const sysClr = getTextByPathList(node, ['a:sysClr', 'attrs', 'lastClr'])
            if (sysClr !== undefined) {
                color = sysClr
            }
        }
        return color
    }

    function toHex(n) {
        let hex = n.toString(16)
        while (hex.length < 2) {
            hex = '0' + hex
        }
        return hex
    }

    function hslToRgb(hue, sat, light) {
        let t1, t2, r, g, b
        hue = hue / 60
        if (light <= 0.5) {
            t2 = light * (sat + 1)
        } else {
            t2 = light + sat - (light * sat)
        }
        t1 = light * 2 - t2
        r = hueToRgb(t1, t2, hue + 2) * 255
        g = hueToRgb(t1, t2, hue) * 255
        b = hueToRgb(t1, t2, hue - 2) * 255
        return {r: r, g: g, b: b}
    }

    function hueToRgb(t1, t2, hue) {
        if (hue < 0) hue += 6
        if (hue >= 6) hue -= 6
        if (hue < 1) return (t2 - t1) * hue + t1
        else if (hue < 3) return t2
        else if (hue < 4) return (t2 - t1) * (4 - hue) + t1
        else return t1
    }

    function getColorName2Hex(name) {
        let hex
        const colorName = ['AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'Black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen']
        const colorHex = ['f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32']
        const findIndx = colorName.indexOf(name)
        if (findIndx !== -1) {
            hex = colorHex[findIndx]
        }
        return hex
    }

    function getColorOpacity(solidFill) {
        if (solidFill === undefined) {
            return undefined
        }
        let opcity = 0

        if (solidFill['a:srgbClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:srgbClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        } else if (solidFill['a:schemeClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:schemeClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        } else if (solidFill['a:scrgbClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:scrgbClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        } else if (solidFill['a:prstClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:prstClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        } else if (solidFill['a:hslClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:hslClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        } else if (solidFill['a:sysClr'] !== undefined) {
            const tint = getTextByPathList(solidFill, ['a:sysClr', 'a:tint', 'attrs', 'val'])
            if (tint !== undefined) {
                opcity = parseInt(tint) / 100000
            }
        }

        return opcity
    }

    function getSchemeColorFromTheme(schemeClr, sldMasterNode) {
        // <p:clrMap ...> in slide master
        // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride

        if (slideLayoutClrOvride === '' || slideLayoutClrOvride === undefined) {
            slideLayoutClrOvride = getTextByPathList(sldMasterNode, ['p:sldMaster', 'p:clrMap', 'attrs'])
        }
        // console.log(slideLayoutClrOvride);
        const schmClrName = schemeClr.substr(2)
        if (slideLayoutClrOvride) {
            switch (schmClrName) {
                case 'tx1':
                case 'tx2':
                case 'bg1':
                case 'bg2': {
                    schemeClr = 'a:' + slideLayoutClrOvride[schmClrName]
                    // console.log(schmClrName+ "=> "+schemeClr);
                    break
                }
            }
        }


        const refNode = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme', schemeClr])
        let color = getTextByPathList(refNode, ['a:srgbClr', 'attrs', 'val'])
        if (color === undefined) {
            color = getTextByPathList(refNode, ['a:sysClr', 'attrs', 'lastClr'])
        }
        return color
    }


// ===== Node functions =====
    /**
     * getTextByPathStr
     * @param {Object} node
     * @param {string} pathStr
     */
    function getTextByPathStr(node, pathStr) {
        return getTextByPathList(node, pathStr.trim().split(/\s+/))
    }

    /**
     * getTextByPathList
     * @param {Object} node
     * @param {Array.<string>} path
     */
    function getTextByPathList(node, path) {
        if (path.constructor !== Array) {
            throw Error('Error of path type! path is not array.')
        }

        if (node === undefined) {
            return undefined
        }

        const l = path.length
        for (let i = 0; i < l; i++) {
            node = node[path[i]]
            if (node === undefined) {
                return undefined
            }
        }

        return node
    }


    function angleToDegrees(angle) {
        if (angle === '' || angle == null) {
            return 0
        }
        return Math.round(angle / 60000)
    }

    function getImageMimeType(imgFileExt) {
        let mimeType = ''
        // console.log(imgFileExt)
        switch (imgFileExt.toLowerCase()) {
            case 'jpg':
            case 'jpeg': {
                mimeType = 'image/jpeg'
                break
            }
            case 'png': {
                mimeType = 'image/png'
                break
            }
            case 'gif': {
                mimeType = 'image/gif'
                break
            }
            case 'emf': { // Not native support
                mimeType = 'image/x-emf'
                break
            }
            case 'wmf': { // Not native support
                mimeType = 'image/x-wmf'
                break
            }
            case 'svg': {
                mimeType = 'image/svg+xml'
                break
            }
            default: {
                mimeType = 'image/*'
            }
        }
        return mimeType
    }


    function getMimeType(imgFileExt) {
        let mimeType = ''
        //console.log(imgFileExt)
        switch (imgFileExt.toLowerCase()) {
            case 'jpg':
            case 'jpeg':
                mimeType = 'image/jpeg'
                break
            case 'png':
                mimeType = 'image/png'
                break
            case 'gif':
                mimeType = 'image/gif'
                break
            case 'emf': // Not native support
                mimeType = 'image/x-emf'
                break
            case 'wmf': // Not native support
                mimeType = 'image/x-wmf'
                break
            case 'svg':
                mimeType = 'image/svg+xml'
                break
            case 'mp4':
                mimeType = 'video/mp4'
                break
            case 'webm':
                mimeType = 'video/webm'
                break
            case 'ogg':
                mimeType = 'video/ogg'
                break
            case 'avi':
                mimeType = 'video/avi'
                break
            case 'mpg':
                mimeType = 'video/mpg'
                break
            case 'wmv':
                mimeType = 'video/wmv'
                break
            case 'mp3':
                mimeType = 'audio/mpeg'
                break
            case 'wav':
                mimeType = 'audio/wav'
                break
        }
        return mimeType
    }
}
