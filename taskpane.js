!function(){"use strict";var t={58394:function(t,e,n){t.exports=n.p+"1fda685b81e1123773f6.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var c=e[r]={exports:{}};return t[r](c,c.exports,n),c.exports}n.m=t,n.d=function(t,e){for(var r in e)n.o(e,r)&&!n.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&"SCRIPT"===e.currentScript.tagName.toUpperCase()&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&(!t||!/^http(s?):/.test(t));)t=r[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,Office.onReady((function(){var t,e=document.getElementById("urlInput"),n=null!==(t=localStorage.getItem("sharedUrl"))&&void 0!==t?t:"";e.value=n,Office.context.document.settings.set("apiUrl",n),Office.context.document.settings.saveAsync()})),document.addEventListener("DOMContentLoaded",(function(){var t=document.getElementById("urlInput");null==t||t.addEventListener("input",(function(){var e=t.value.trim();e&&(Office.context.document.settings.set("apiUrl",e),Office.context.document.settings.saveAsync(),localStorage.setItem("sharedUrl",e))}))})),new URL(n(58394),n.b)}();
//# sourceMappingURL=taskpane.js.map