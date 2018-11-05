export default function (resultElement) {

    return `${resultElement} section {
	width: 100%;
	height: 690px;
	position: relative;
	text-align: center;
	overflow: hidden;
}

${resultElement} section div.block {
	position: absolute;
	top: 0px;
	left: 0px;
	width: 100%;
}

${resultElement} section div.content {
  display: flex;
  flex-direction: column;
  /*
  justify-content: center;
  align-items: flex-end;
  */
}

${resultElement} section div.v-up {
  justify-content: flex-start;
}
${resultElement} section div.v-mid {
  justify-content: center;
}
${resultElement} section div.v-down {
  justify-content: flex-end;
}
${resultElement} section div.h-left {
  align-items: flex-start;
  text-align: left;
}
${resultElement}section div.h-mid {
  align-items: center;
  text-align: center;
}
${resultElement} section div.h-right {
  align-items: flex-end;
  text-align: right;
}
${resultElement} section div.up-left {
  justify-content: flex-start;
  align-items: flex-start;
  text-align: left;
}
${resultElement} section div.up-center {
  justify-content: flex-start;
  align-items: center;
}
${resultElement} section div.up-right {
  justify-content: flex-start;
  align-items: flex-end;
}
${resultElement} section div.center-left {
  justify-content: center;
  align-items: flex-start;
  text-align: left;
}
${resultElement} section div.center-center {
  justify-content: center;
  align-items: center;
}
${resultElement} section div.center-right {
  justify-content: center;
  align-items: flex-end;
}
${resultElement} section div.down-left {
  justify-content: flex-end;
  align-items: flex-start;
  text-align: left;
}
${resultElement} section div.down-center {
  justify-content: flex-end;
  align-items: center;
}
${resultElement} section div.down-right {
  justify-content: flex-end;
  align-items: flex-end;
}

${resultElement} section span.text-block {
  /* display: inline-block; */
}

${resultElement} li.slide {
  margin: 10px 0;
  font-size: 18px;
}

${resultElement} div.footer {
  text-align: center;
}

${resultElement} section table {
  position: absolute;
}

${resultElement} section table, section th, section td {
  border: 1px solid black;
}

${resultElement} section svg.drawing {
  position: absolute;
  overflow: visible;
}`

}
