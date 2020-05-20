'use strict'

import $ from 'jquery'
import dimple from 'dimple'
import processPptx from './process_pptx'
import pptxStyle from './pptx_css'

/**
 * @param {ArrayBuffer} pptx
 * @param {Element|String} resultElement
 * @param {Element|String} [thumbElement]
 */
const renderPptx = (pptx, resultElement, thumbElement) => {
    const $result = $(resultElement)
    const $wrapper = $('<div class="pptx-wrapper"></div>')
    $result.html('')
    $result.append($wrapper)
    // $wrapper.append(`<style>${pptxStyle(resultElement)}</style>`)
    let isDone = false

    return new Promise((resolve, reject) => {
        const processMessage = (msg) => {
            if (isDone) return
            switch (msg.type) {
                case 'slide':
                    $wrapper.append(msg.data)
                    break
                case 'pptx-thumb':
                    // if (thumbElement) $(thumbElement).attr('src', `data:image/jpeg;base64,${msg.data}`)
                    break
                case 'slideSize':
                    break
                case 'globalCSS':
                    $wrapper.append(`<style>${msg.data}</style>`)
                    break
                case 'Done':
                    isDone = true
                    resolve({time: msg.data.time, slideSize: msg.data.slideSize})
                    break
                case 'WARN':
                    console.warn('PPTX processing warning: ', msg.data)
                    break
                case 'ERROR':
                    isDone = true
                    console.error('PPTX processing error: ', msg.data)
                    reject(new Error(msg.data))
                    break
                case 'DEBUG':
                     console.debug('Worker: ', msg.data);
                    break
                case 'INFO':
                default:
                 console.info('Worker: ', msg.data);
            }
        }
        /*
        // Actual Web Worker - If you want to use this, switching worker's url to Blob is probably better
        const worker = new Worker('./dist/worker.js')
        worker.addEventListener('message', event => processMessage(event.data), false)
        const stopWorker = setInterval(() => { // Maybe this should be done in the message processing
          if (isDone) {
            worker.terminate()
            // console.log("worker terminated");
            clearInterval(stopWorker)
          }
        }, 500)
        */
        const worker = { // shim worker
            postMessage: () => {
            },
            terminate: () => {
            }
        }
        processPptx(
            func => {
                worker.postMessage = func
            },
            processMessage
        )
        worker.postMessage({
            'type': 'processPPTX',
            'data': pptx
        })
    }).then(time => {
        const resize = () => {
            const slidesWidth = Math.max(...Array.from($wrapper.children('section')).map(s => s.offsetWidth))
            const wrapperWidth = $wrapper[0].offsetWidth
            $wrapper.css({
                'transform': `scale(${wrapperWidth / slidesWidth})`,
                'transform-origin': 'top left'
            })
        }
        resize()
        window.addEventListener('resize', resize)
        setNumericBullets($('.block'))
        setNumericBullets($('table td'))
        return time
    })
}

export default renderPptx


function setNumericBullets(elem) {
    const paragraphsArray = elem
    for (let i = 0; i < paragraphsArray.length; i++) {
        const buSpan = $(paragraphsArray[i]).find('.numeric-bullet-style')
        if (buSpan.length > 0) {
            // console.log("DIV-"+i+":");
            let prevBultTyp = ''
            let prevBultLvl = ''
            let buletIndex = 0
            const tmpArry = []
            let tmpArryIndx = 0
            const buletTypSrry = []
            for (let j = 0; j < buSpan.length; j++) {
                const bulletType = $(buSpan[j]).data('bulltname')
                const bulletLvl = $(buSpan[j]).data('bulltlvl')
                // console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
                if (buletIndex === 0) {
                    prevBultTyp = bulletType
                    prevBultLvl = bulletLvl
                    tmpArry[tmpArryIndx] = buletIndex
                    buletTypSrry[tmpArryIndx] = bulletType
                    buletIndex++
                } else {
                    if (bulletType === prevBultTyp && bulletLvl === prevBultLvl) {
                        prevBultTyp = bulletType
                        prevBultLvl = bulletLvl
                        buletIndex++
                        tmpArry[tmpArryIndx] = buletIndex
                        buletTypSrry[tmpArryIndx] = bulletType
                    } else if (bulletType !== prevBultTyp && bulletLvl === prevBultLvl) {
                        prevBultTyp = bulletType
                        prevBultLvl = bulletLvl
                        tmpArryIndx++
                        tmpArry[tmpArryIndx] = buletIndex
                        buletTypSrry[tmpArryIndx] = bulletType
                        buletIndex = 1
                    } else if (bulletType !== prevBultTyp && Number(bulletLvl) > Number(prevBultLvl)) {
                        prevBultTyp = bulletType
                        prevBultLvl = bulletLvl
                        tmpArryIndx++
                        tmpArry[tmpArryIndx] = buletIndex
                        buletTypSrry[tmpArryIndx] = bulletType
                        buletIndex = 1
                    } else if (bulletType !== prevBultTyp && Number(bulletLvl) < Number(prevBultLvl)) {
                        prevBultTyp = bulletType
                        prevBultLvl = bulletLvl
                        tmpArryIndx--
                        buletIndex = tmpArry[tmpArryIndx] + 1
                    }
                }
                // console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
                const numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex)
                $(buSpan[j]).html(numIdx)
            }
        }
    }
}

function getNumTypeNum(numTyp, num) {
    let rtrnNum = ''
    switch (numTyp) {
        case 'arabicPeriod':
            rtrnNum = num + '. '
            break
        case 'arabicParenR':
            rtrnNum = num + ') '
            break
        case 'alphaLcParenR':
            rtrnNum = alphaNumeric(num, 'lowerCase') + ') '
            break
        case 'alphaLcPeriod':
            rtrnNum = alphaNumeric(num, 'lowerCase') + '. '
            break

        case 'alphaUcParenR':
            rtrnNum = alphaNumeric(num, 'upperCase') + ') '
            break
        case 'alphaUcPeriod':
            rtrnNum = alphaNumeric(num, 'upperCase') + '. '
            break

        case 'romanUcPeriod':
            rtrnNum = romanize(num) + '. '
            break
        case 'romanLcParenR':
            rtrnNum = romanize(num) + ') '
            break
        case 'hebrew2Minus':
            rtrnNum = hebrew2Minus.format(num) + '-'
            break
        default:
            rtrnNum = num
    }
    return rtrnNum
}

function romanize(num) {
    if (!+num) return false
    const digits = String(+num).split('')
    const key = ['', 'C', 'CC', 'CCC', 'CD', 'D', 'DC', 'DCC', 'DCCC', 'CM',
        '', 'X', 'XX', 'XXX', 'XL', 'L', 'LX', 'LXX', 'LXXX', 'XC',
        '', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX']
    let roman = ''
    let i = 3
    while (i--) roman = (key[+digits.pop() + (i * 10)] || '') + roman
    return (new Array(+digits.join('') + 1)).join('M') + roman
}

const hebrew2Minus = archaicNumbers([
    [1000, ''],
    [400, 'ת'],
    [300, 'ש'],
    [200, 'ר'],
    [100, 'ק'],
    [90, 'צ'],
    [80, 'פ'],
    [70, 'ע'],
    [60, 'ס'],
    [50, 'נ'],
    [40, 'מ'],
    [30, 'ל'],
    [20, 'כ'],
    [10, 'י'],
    [9, 'ט'],
    [8, 'ח'],
    [7, 'ז'],
    [6, 'ו'],
    [5, 'ה'],
    [4, 'ד'],
    [3, 'ג'],
    [2, 'ב'],
    [1, 'א'],
    [/יה/, 'ט״ו'],
    [/יו/, 'ט״ז'],
    [/([א-ת])([א-ת])$/, '$1״$2'],
    [/^([א-ת])$/, '$1׳']
])

function archaicNumbers(arr) {
    // const arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length })
    return {
        format: function (n) {
            let ret = ''
            $.each(arr, function () {
                const num = this[0]
                if (parseInt(num) > 0) {
                    for (; n >= num; n -= num) ret += this[1]
                } else {
                    ret = ret.replace(num, this[1])
                }
            })
            return ret
        }
    }
}

function alphaNumeric(num, upperLower) {
    num = Number(num) - 1
    let aNum = ''
    if (upperLower === 'upperCase') {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase()
    } else if (upperLower === 'lowerCase') {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase()
    }
    return aNum
}
