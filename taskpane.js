!function(){"use strict";var t={58394:function(t,e,r){t.exports=r.p+"1fda685b81e1123773f6.css"}},e={};function r(n){var o=e[n];if(void 0!==o)return o.exports;var c=e[n]={exports:{}};return t[n](c,c.exports,r),c.exports}r.m=t,r.d=function(t,e){for(var n in e)r.o(e,n)&&!r.o(t,n)&&Object.defineProperty(t,n,{enumerable:!0,get:e[n]})},r.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),r.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;r.g.importScripts&&(t=r.g.location+"");var e=r.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var n=e.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=n[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),r.p=t}(),r.b=document.baseURI||self.location.href,Office.onReady((function(){var t;document.getElementById("urlInput").value=null!==(t=localStorage.getItem("sharedUrl"))&&void 0!==t?t:""})),document.addEventListener("DOMContentLoaded",(function(){var t=document.getElementById("urlInput");null==t||t.addEventListener("input",(function(){var e=t.value.trim();e&&(Office.context.document.settings.set("apiUrl",e),Office.context.document.settings.saveAsync(),localStorage.setItem("sharedUrl",e))}))})),new URL(r(58394),r.b)}();
//# sourceMappingURL=taskpane.js.map