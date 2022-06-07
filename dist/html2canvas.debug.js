/*!
  html2canvas  <http://html2canvas.hertzen.com>
  Copyright (c) 2016 Niklas von Hertzen
*/

/*! Copyright Mathias Bynens <https://mathiasbynens.be/> */
var PunyCode;
(function (PunyCode_1) {
	var maxInt = 2147483647, base = 36, tMin = 1, tMax = 26, skew = 38, damp = 700, initialBias = 72, initialN = 128, delimiter = '-', regexPunycode = /^xn--/, regexNonASCII = /[^\x20-\x7E]/, regexSeparators = /[\x2E\u3002\uFF0E\uFF61]/g, errors = {
		'overflow': 'Overflow: input needs wider integers to process',
		'not-basic': 'Illegal input >= 0x80 (not a basic code point)',
		'invalid-input': 'Invalid input'
	}, baseMinusTMin = base - tMin, floor = Math.floor, stringFromCharCode = String.fromCharCode, key;
	function error(type) {
		throw RangeError(errors[type]);
	}
	function map(array, fn) {
		var length = array.length;
		var result = [];
		while (length--) {
			result[length] = fn(array[length]);
		}
		return result;
	}
	function mapDomain(string, fn) {
		var parts = string.split('@');
		var result = '';
		if (parts.length > 1) {
			result = parts[0] + '@';
			string = parts[1];
		}
		string = string.replace(regexSeparators, '\x2E');
		var labels = string.split('.');
		var encoded = map(labels, fn).join('.');
		return result + encoded;
	}
	function ucs2decode(string) {
		var output = [], counter = 0, length = string.length, value, extra;
		while (counter < length) {
			value = string.charCodeAt(counter++);
			if (value >= 0xD800 && value <= 0xDBFF && counter < length) {
				extra = string.charCodeAt(counter++);
				if ((extra & 0xFC00) == 0xDC00) {
					output.push(((value & 0x3FF) << 10) + (extra & 0x3FF) + 0x10000);
				}
				else {
					output.push(value);
					counter--;
				}
			}
			else {
				output.push(value);
			}
		}
		return output;
	}
	function ucs2encode(array) {
		return map(array, function (value) {
			var output = '';
			if (value > 0xFFFF) {
				value -= 0x10000;
				output += stringFromCharCode(value >>> 10 & 0x3FF | 0xD800);
				value = 0xDC00 | value & 0x3FF;
			}
			output += stringFromCharCode(value);
			return output;
		}).join('');
	}
	function basicToDigit(codePoint) {
		if (codePoint - 48 < 10) {
			return codePoint - 22;
		}
		if (codePoint - 65 < 26) {
			return codePoint - 65;
		}
		if (codePoint - 97 < 26) {
			return codePoint - 97;
		}
		return base;
	}
	function digitToBasic(digit, flag) {
		return digit + 22 + 75 * (Number)(digit < 26) - ((Number)(flag != 0) << 5);
	}
	function adapt(delta, numPoints, firstTime) {
		var k = 0;
		delta = firstTime ? floor(delta / damp) : delta >> 1;
		delta += floor(delta / numPoints);
		for (; delta > baseMinusTMin * tMax >> 1; k += base) {
			delta = floor(delta / baseMinusTMin);
		}
		return floor(k + (baseMinusTMin + 1) * delta / (delta + skew));
	}
	function decode(input) {
		var output = [], inputLength = input.length, out, i = 0, n = initialN, bias = initialBias, basic, j, index, oldi, w, k, digit, t, baseMinusT;
		basic = input.lastIndexOf(delimiter);
		if (basic < 0) {
			basic = 0;
		}
		for (j = 0; j < basic; ++j) {
			if (input.charCodeAt(j) >= 0x80) {
				error('not-basic');
			}
			output.push(input.charCodeAt(j));
		}
		for (index = basic > 0 ? basic + 1 : 0; index < inputLength;) {
			for (oldi = i, w = 1, k = base; ; k += base) {
				if (index >= inputLength) {
					error('invalid-input');
				}
				digit = basicToDigit(input.charCodeAt(index++));
				if (digit >= base || digit > floor((maxInt - i) / w)) {
					error('overflow');
				}
				i += digit * w;
				t = k <= bias ? tMin : (k >= bias + tMax ? tMax : k - bias);
				if (digit < t) {
					break;
				}
				baseMinusT = base - t;
				if (w > floor(maxInt / baseMinusT)) {
					error('overflow');
				}
				w *= baseMinusT;
			}
			out = output.length + 1;
			bias = adapt(i - oldi, out, oldi == 0);
			if (floor(i / out) > maxInt - n) {
				error('overflow');
			}
			n += floor(i / out);
			i %= out;
			output.splice(i++, 0, n);
		}
		return ucs2encode(output);
	}
	function encode(input) {
		var n, delta, handledCPCount, basicLength, bias, j, m, q, k, t, currentValue, output = [], inputLength, handledCPCountPlusOne, baseMinusT, qMinusT;
		input = ucs2decode(input);
		inputLength = input.length;
		n = initialN;
		delta = 0;
		bias = initialBias;
		for (j = 0; j < inputLength; ++j) {
			currentValue = input[j];
			if (currentValue < 0x80) {
				output.push(stringFromCharCode(currentValue));
			}
		}
		handledCPCount = basicLength = output.length;
		if (basicLength) {
			output.push(delimiter);
		}
		while (handledCPCount < inputLength) {
			for (m = maxInt, j = 0; j < inputLength; ++j) {
				currentValue = input[j];
				if (currentValue >= n && currentValue < m) {
					m = currentValue;
				}
			}
			handledCPCountPlusOne = handledCPCount + 1;
			if (m - n > floor((maxInt - delta) / handledCPCountPlusOne)) {
				error('overflow');
			}
			delta += (m - n) * handledCPCountPlusOne;
			n = m;
			for (j = 0; j < inputLength; ++j) {
				currentValue = input[j];
				if (currentValue < n && ++delta > maxInt) {
					error('overflow');
				}
				if (currentValue == n) {
					for (q = delta, k = base; ; k += base) {
						t = k <= bias ? tMin : (k >= bias + tMax ? tMax : k - bias);
						if (q < t) {
							break;
						}
						qMinusT = q - t;
						baseMinusT = base - t;
						output.push(stringFromCharCode(digitToBasic(t + qMinusT % baseMinusT, 0)));
						q = floor(qMinusT / baseMinusT);
					}
					output.push(stringFromCharCode(digitToBasic(q, 0)));
					bias = adapt(delta, handledCPCountPlusOne, handledCPCount == basicLength);
					delta = 0;
					++handledCPCount;
				}
			}
			++delta;
			++n;
		}
		return output.join('');
	}
	function toUnicode(input) {
		return mapDomain(input, function (string) {
			return regexPunycode.test(string)
                ? decode(string.slice(4).toLowerCase())
                : string;
		});
	}
	function toASCII(input) {
		return mapDomain(input, function (string) {
			return regexNonASCII.test(string)
                ? 'xn--' + encode(string)
                : string;
		});
	}
	PunyCode_1.PunyCode = {
		'version': '1.3.2',
		'ucs2': {
			'decode': ucs2decode,
			'encode': ucs2encode
		},
		'decode': decode,
		'encode': encode,
		'toASCII': toASCII,
		'toUnicode': toUnicode
	};
})(PunyCode || (PunyCode = {}));
var Html2canvas;
(function (Html2canvas) {
	Html2canvas.Color = function (value) {
		this.r = 0;
		this.g = 0;
		this.b = 0;
		this.a = null;
		var result = this.fromArray(value) ||
            this.namedColor(value) ||
            this.rgb(value) ||
            this.rgba(value) ||
            this.hex6(value) ||
            this.hex3(value);
	};
	Html2canvas.Color.prototype.darken = function (amount) {
		var a = 1 - amount;
		return new Html2canvas.Color([
            Math.round(this.r * a),
            Math.round(this.g * a),
            Math.round(this.b * a),
            this.a
		]);
	};
	Html2canvas.Color.prototype.isTransparent = function () {
		return this.a === 0;
	};
	Html2canvas.Color.prototype.isBlack = function () {
		return this.r === 0 && this.g === 0 && this.b === 0;
	};
	Html2canvas.Color.prototype.fromArray = function (array) {
		if (Array.isArray(array)) {
			this.r = Math.min(array[0], 255);
			this.g = Math.min(array[1], 255);
			this.b = Math.min(array[2], 255);
			if (array.length > 3) {
				this.a = array[3];
			}
		}
		return (Array.isArray(array));
	};
	var _hex3 = /^#([a-f0-9]{3})$/i;
	Html2canvas.Color.prototype.hex3 = function (value) {
		var match = null;
		if ((match = value.match(_hex3)) !== null) {
			this.r = parseInt(match[1][0] + match[1][0], 16);
			this.g = parseInt(match[1][1] + match[1][1], 16);
			this.b = parseInt(match[1][2] + match[1][2], 16);
		}
		return match !== null;
	};
	var _hex6 = /^#([a-f0-9]{6})$/i;
	Html2canvas.Color.prototype.hex6 = function (value) {
		var match = null;
		if ((match = value.match(_hex6)) !== null) {
			this.r = parseInt(match[1].substring(0, 2), 16);
			this.g = parseInt(match[1].substring(2, 4), 16);
			this.b = parseInt(match[1].substring(4, 6), 16);
		}
		return match !== null;
	};
	var _rgb = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/;
	Html2canvas.Color.prototype.rgb = function (value) {
		var match = null;
		if ((match = value.match(_rgb)) !== null) {
			this.r = Number(match[1]);
			this.g = Number(match[2]);
			this.b = Number(match[3]);
		}
		return match !== null;
	};
	var _rgba = /^rgba\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d?\.?\d+)\s*\)$/;
	Html2canvas.Color.prototype.rgba = function (value) {
		var match = null;
		if ((match = value.match(_rgba)) !== null) {
			this.r = Number(match[1]);
			this.g = Number(match[2]);
			this.b = Number(match[3]);
			this.a = Number(match[4]);
		}
		return match !== null;
	};
	Html2canvas.Color.prototype.toString = function () {
		return this.a !== null && this.a !== 1 ?
            "rgba(" + [this.r, this.g, this.b, this.a].join(",") + ")" :
            "rgb(" + [this.r, this.g, this.b].join(",") + ")";
	};
	Html2canvas.Color.prototype.namedColor = function (value) {
		value = value.toLowerCase();
		var color = colors[value];
		if (color) {
			this.r = color[0];
			this.g = color[1];
			this.b = color[2];
		}
		else if (value === "transparent") {
			this.r = this.g = this.b = this.a = 0;
			return true;
		}
		return !!color;
	};
	Html2canvas.Color.prototype.isColor = true;
	var colors = {
		"aliceblue": [240, 248, 255],
		"antiquewhite": [250, 235, 215],
		"aqua": [0, 255, 255],
		"aquamarine": [127, 255, 212],
		"azure": [240, 255, 255],
		"beige": [245, 245, 220],
		"bisque": [255, 228, 196],
		"black": [0, 0, 0],
		"blanchedalmond": [255, 235, 205],
		"blue": [0, 0, 255],
		"blueviolet": [138, 43, 226],
		"brown": [165, 42, 42],
		"burlywood": [222, 184, 135],
		"cadetblue": [95, 158, 160],
		"chartreuse": [127, 255, 0],
		"chocolate": [210, 105, 30],
		"coral": [255, 127, 80],
		"cornflowerblue": [100, 149, 237],
		"cornsilk": [255, 248, 220],
		"crimson": [220, 20, 60],
		"cyan": [0, 255, 255],
		"darkblue": [0, 0, 139],
		"darkcyan": [0, 139, 139],
		"darkgoldenrod": [184, 134, 11],
		"darkgray": [169, 169, 169],
		"darkgreen": [0, 100, 0],
		"darkgrey": [169, 169, 169],
		"darkkhaki": [189, 183, 107],
		"darkmagenta": [139, 0, 139],
		"darkolivegreen": [85, 107, 47],
		"darkorange": [255, 140, 0],
		"darkorchid": [153, 50, 204],
		"darkred": [139, 0, 0],
		"darksalmon": [233, 150, 122],
		"darkseagreen": [143, 188, 143],
		"darkslateblue": [72, 61, 139],
		"darkslategray": [47, 79, 79],
		"darkslategrey": [47, 79, 79],
		"darkturquoise": [0, 206, 209],
		"darkviolet": [148, 0, 211],
		"deeppink": [255, 20, 147],
		"deepskyblue": [0, 191, 255],
		"dimgray": [105, 105, 105],
		"dimgrey": [105, 105, 105],
		"dodgerblue": [30, 144, 255],
		"firebrick": [178, 34, 34],
		"floralwhite": [255, 250, 240],
		"forestgreen": [34, 139, 34],
		"fuchsia": [255, 0, 255],
		"gainsboro": [220, 220, 220],
		"ghostwhite": [248, 248, 255],
		"gold": [255, 215, 0],
		"goldenrod": [218, 165, 32],
		"gray": [128, 128, 128],
		"green": [0, 128, 0],
		"greenyellow": [173, 255, 47],
		"grey": [128, 128, 128],
		"honeydew": [240, 255, 240],
		"hotpink": [255, 105, 180],
		"indianred": [205, 92, 92],
		"indigo": [75, 0, 130],
		"ivory": [255, 255, 240],
		"khaki": [240, 230, 140],
		"lavender": [230, 230, 250],
		"lavenderblush": [255, 240, 245],
		"lawngreen": [124, 252, 0],
		"lemonchiffon": [255, 250, 205],
		"lightblue": [173, 216, 230],
		"lightcoral": [240, 128, 128],
		"lightcyan": [224, 255, 255],
		"lightgoldenrodyellow": [250, 250, 210],
		"lightgray": [211, 211, 211],
		"lightgreen": [144, 238, 144],
		"lightgrey": [211, 211, 211],
		"lightpink": [255, 182, 193],
		"lightsalmon": [255, 160, 122],
		"lightseagreen": [32, 178, 170],
		"lightskyblue": [135, 206, 250],
		"lightslategray": [119, 136, 153],
		"lightslategrey": [119, 136, 153],
		"lightsteelblue": [176, 196, 222],
		"lightyellow": [255, 255, 224],
		"lime": [0, 255, 0],
		"limegreen": [50, 205, 50],
		"linen": [250, 240, 230],
		"magenta": [255, 0, 255],
		"maroon": [128, 0, 0],
		"mediumaquamarine": [102, 205, 170],
		"mediumblue": [0, 0, 205],
		"mediumorchid": [186, 85, 211],
		"mediumpurple": [147, 112, 219],
		"mediumseagreen": [60, 179, 113],
		"mediumslateblue": [123, 104, 238],
		"mediumspringgreen": [0, 250, 154],
		"mediumturquoise": [72, 209, 204],
		"mediumvioletred": [199, 21, 133],
		"midnightblue": [25, 25, 112],
		"mintcream": [245, 255, 250],
		"mistyrose": [255, 228, 225],
		"moccasin": [255, 228, 181],
		"navajowhite": [255, 222, 173],
		"navy": [0, 0, 128],
		"oldlace": [253, 245, 230],
		"olive": [128, 128, 0],
		"olivedrab": [107, 142, 35],
		"orange": [255, 165, 0],
		"orangered": [255, 69, 0],
		"orchid": [218, 112, 214],
		"palegoldenrod": [238, 232, 170],
		"palegreen": [152, 251, 152],
		"paleturquoise": [175, 238, 238],
		"palevioletred": [219, 112, 147],
		"papayawhip": [255, 239, 213],
		"peachpuff": [255, 218, 185],
		"peru": [205, 133, 63],
		"pink": [255, 192, 203],
		"plum": [221, 160, 221],
		"powderblue": [176, 224, 230],
		"purple": [128, 0, 128],
		"rebeccapurple": [102, 51, 153],
		"red": [255, 0, 0],
		"rosybrown": [188, 143, 143],
		"royalblue": [65, 105, 225],
		"saddlebrown": [139, 69, 19],
		"salmon": [250, 128, 114],
		"sandybrown": [244, 164, 96],
		"seagreen": [46, 139, 87],
		"seashell": [255, 245, 238],
		"sienna": [160, 82, 45],
		"silver": [192, 192, 192],
		"skyblue": [135, 206, 235],
		"slateblue": [106, 90, 205],
		"slategray": [112, 128, 144],
		"slategrey": [112, 128, 144],
		"snow": [255, 250, 250],
		"springgreen": [0, 255, 127],
		"steelblue": [70, 130, 180],
		"tan": [210, 180, 140],
		"teal": [0, 128, 128],
		"thistle": [216, 191, 216],
		"tomato": [255, 99, 71],
		"turquoise": [64, 224, 208],
		"violet": [238, 130, 238],
		"wheat": [245, 222, 179],
		"white": [255, 255, 255],
		"whitesmoke": [245, 245, 245],
		"yellow": [255, 255, 0],
		"yellowgreen": [154, 205, 50]
	};
})(Html2canvas || (Html2canvas = {}));
var Html2canvas;
(function (Html2canvas) {
	Html2canvas.Log = function () {
		if (Html2canvas.Log.options.logging && window.console && window.console.log) {
			Function.prototype.bind.call(window.console.log, (window.console)).apply(window.console, [(Date.now() - Html2canvas.Log.options.start) + "ms", "html2canvas:"].concat([].slice.call(arguments, 0)));
		}
	};
	Html2canvas.Log.options = { logging: false };
})(Html2canvas || (Html2canvas = {}));
var Html2canvas;
(function (Html2canvas) {
	Html2canvas.Support = function (document) {
		this.rangeBounds = this.testRangeBounds(document);
		this.cors = this.testCORS();
		this.svg = this.testSVG();
	};
	Html2canvas.Support.prototype.testRangeBounds = function (document) {
		var range, testElement, rangeBounds, rangeHeight, support = false;
		if (document.createRange) {
			range = document.createRange();
			if (range.getBoundingClientRect) {
				testElement = document.createElement('boundtest');
				testElement.style.height = "123px";
				testElement.style.display = "block";
				document.body.appendChild(testElement);
				range.selectNode(testElement);
				rangeBounds = range.getBoundingClientRect();
				rangeHeight = rangeBounds.height;
				if (rangeHeight === 123) {
					support = true;
				}
				document.body.removeChild(testElement);
			}
		}
		return support;
	};
	Html2canvas.Support.prototype.testCORS = function () {
		return typeof ((new Image()).crossOrigin) !== "undefined";
	};
	Html2canvas.Support.prototype.testSVG = function () {
		var img = new