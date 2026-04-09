/**
 * src/utils/path-browser-shim.js
 *
 * kuromoji の DictionaryLoader が require("path") で呼ぶ path.join を
 * ブラウザ環境でも動くよう最小限に実装したシム。
 *
 * 通常の posix join に加えて HTTP(S) URL をそのまま結合できる。
 * （path-browserify は https:// の二重スラッシュを正規化して壊すため使えない）
 */

function join() {
  var parts = Array.prototype.slice.call(arguments)
  var str = parts.join('/')
  // https:// または http:// のプロトコル部分を保護してから正規化
  var match = str.match(/^(https?:\/\/)(.*)$/)
  if (match) {
    return match[1] + match[2].replace(/\/+/g, '/')
  }
  return str.replace(/\/+/g, '/')
}

module.exports = { join: join }
