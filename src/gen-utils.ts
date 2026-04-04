/**
 * PptxGenJS: Utility Methods
 */

import { EMU, REGEX_HEX_COLOR, DEF_FONT_COLOR, ONEPT, SchemeColor, SCHEME_COLORS } from './core-enums'
import { PresLayout, TextGlowProps, PresSlide, ShapeFillProps, Color, ShapeLineProps, Coord, ShadowProps } from './core-interfaces'

/**
 * Translates any type of `x`/`y`/`w`/`h` prop to EMU
 * - guaranteed to return a result regardless of undefined, null, etc. (0)
 * - {number} - 12800 (EMU)
 * - {number} - 0.5 (inches)
 * - {string} - "75%"
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber (size: Coord, xyDir: 'X' | 'Y', layout: PresLayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// CASE 1: Number in inches
	// Assume any number less than 100 is inches
	if (typeof size === 'number' && size < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number greater than 100 sure isnt inches! Just return it (assume value is EMU already).
	if (typeof size === 'number' && size >= 100) return size

	// CASE 3: Percentage (ex: '50%')
	if (typeof size === 'string' && size.includes('%')) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid (uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		const r = (Math.random() * 16) | 0
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities (xml: string): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu (inches: number | string): number {
	// NOTE: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, so lets assume its in EMU already, therefore, just return the same value
	if (typeof inches === 'number' && inches > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts (pt: number | string): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees (d: number): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * 60000)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex (c: number): string {
	const hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex (r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**  TODO: FUTURE: TODO-4.0:
 * @date 2022-04-10
 * @tldr this s/b a private method with all current calls switched to `genXmlColorSelection()`
 * @desc lots of code calls this method
 * @example [gen-charts.tx] `strXml += '<a:solidFill>' + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1000)}"/>`) + '</a:solidFill>'`
 * Thi sis wrong. We s/b calling `genXmlColorSelection()` instead as it returns `<a:solidfill>BLAH</a:solidFill>`!!
 */
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export function createColorElement (colorStr: string | SCHEME_COLORS, innerElements?: string): string {
	let colorVal = (colorStr || '').replace('#', '')

	if (
		!REGEX_HEX_COLOR.test(colorVal) &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`)
		colorVal = DEF_FONT_COLOR
	}

	const tagName = REGEX_HEX_COLOR.test(colorVal) ? 'srgbClr' : 'schemeClr'
	const colorAttr = 'val="' + (REGEX_HEX_COLOR.test(colorVal) ? colorVal.toUpperCase() : colorVal) + '"'

	return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', transparency: 25 };
 */
export function createGlowElement (options: TextGlowProps, defaults: TextGlowProps): string {
	let strXml = ''
	const opts = { ...defaults, ...options }
	const size = Math.round(opts.size * ONEPT)
	const color = opts.color

	// Support both old 'opacity' (0-1) and new 'transparency' (0-100) properties
	let alphaVal: number
	if (opts.transparency !== undefined) {
		alphaVal = Math.round((100 - opts.transparency) * 1000)
	} else if ((opts as any).opacity !== undefined) {
		// @deprecated: support old opacity property
		alphaVal = Math.round((opts as any).opacity * 100000)
	} else {
		alphaVal = 100000 // fully opaque by default
	}

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${alphaVal}"/>`)
	strXml += '</a:glow>'

	return strXml
}

/**
 * Generate gradient fill XML
 * @param {import('./core-interfaces').GradientFillProps} gradient - gradient properties
 * @returns {string} XML string
 */
export function genXmlGradientFill (gradient: any): string {
	if (!gradient || !gradient.stops || gradient.stops.length < 2) {
		console.warn('Gradient fill requires at least 2 color stops')
		return ''
	}

	let xml = '<a:gradFill'
	if (gradient.rotateWithShape !== false) xml += ' rotWithShape="1"'
	xml += '>'

	// Add gradient stops
	xml += '<a:gsLst>'
	gradient.stops.forEach((stop: any) => {
		const pos = Math.round(stop.position * 1000) // Convert 0-100 to 0-100000
		xml += `<a:gs pos="${pos}">`

		let innerElements = ''
		if (stop.transparency) {
			innerElements += `<a:alpha val="${Math.round((100 - stop.transparency) * 1000)}"/>`
		}
		xml += createColorElement(stop.color, innerElements)
		xml += '</a:gs>'
	})
	xml += '</a:gsLst>'

	// Add gradient direction/type
	const angle = gradient.angle || 0
	switch (gradient.type) {
		case 'linear':
			xml += `<a:lin ang="${convertRotationDegrees(angle)}" scaled="0"/>`
			break
		case 'radial':
			xml += '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
			break
		case 'rectangular':
			xml += '<a:path path="rect"><a:fillToRect l="0" t="0" r="0" b="0"/></a:path>'
			break
		case 'path':
			xml += '<a:path path="shape"><a:fillToRect l="0" t="0" r="0" b="0"/></a:path>'
			break
		default:
			xml += `<a:lin ang="${convertRotationDegrees(angle)}" scaled="0"/>`
	}

	xml += '</a:gradFill>'
	return xml
}

/**
 * Generate pattern fill XML
 * @param {import('./core-interfaces').PatternFillProps} pattern - pattern properties
 * @returns {string} XML string
 */
export function genXmlPatternFill (pattern: any): string {
	if (!pattern || !pattern.type || !pattern.fgColor || !pattern.bgColor) {
		console.warn('Pattern fill requires type, fgColor, and bgColor')
		return ''
	}

	let xml = `<a:pattFill prst="${pattern.type}">`

	// Foreground color
	xml += '<a:fgClr>'
	let fgInner = ''
	if (pattern.fgTransparency) {
		fgInner += `<a:alpha val="${Math.round((100 - pattern.fgTransparency) * 1000)}"/>`
	}
	xml += createColorElement(pattern.fgColor, fgInner)
	xml += '</a:fgClr>'

	// Background color
	xml += '<a:bgClr>'
	let bgInner = ''
	if (pattern.bgTransparency) {
		bgInner += `<a:alpha val="${Math.round((100 - pattern.bgTransparency) * 1000)}"/>`
	}
	xml += createColorElement(pattern.bgColor, bgInner)
	xml += '</a:bgClr>'

	xml += '</a:pattFill>'
	return xml
}

/**
 * Generate texture/image fill XML
 * @param {import('./core-interfaces').TextureFillProps} texture - texture properties
 * @param {number} rId - relationship ID for the image
 * @returns {string} XML string
 */
export function genXmlTextureFill (texture: any, rId: number): string {
	if (!texture || (!texture.path && !texture.data)) {
		console.warn('Texture fill requires path or data')
		return ''
	}

	let xml = '<a:blipFill'
	if (texture.tile !== false) xml += ' dpi="0" rotWithShape="1"'
	xml += '>'

	xml += `<a:blip r:embed="rId${rId}">`
	if (texture.transparency) {
		xml += `<a:alphaModFix amt="${Math.round((100 - texture.transparency) * 1000)}"/>`
	}
	xml += '</a:blip>'

	if (texture.tile !== false) {
		xml += '<a:tile/>'
	} else {
		// Stretch mode with alignment
		xml += '<a:stretch><a:fillRect'
		if (texture.alignX === 'left') xml += ' l="0" r="50000"'
		else if (texture.alignX === 'right') xml += ' l="50000" r="0"'
		if (texture.alignY === 'top') xml += ' t="0" b="50000"'
		else if (texture.alignY === 'bottom') xml += ' t="50000" b="0"'
		xml += '/></a:stretch>'
	}

	xml += '</a:blipFill>'
	return xml
}

/**
 * Generate reflection effect XML
 * @param {import('./core-interfaces').ReflectionProps} reflection - reflection properties
 * @returns {string} XML string
 */
export function genXmlReflection (reflection: any): string {
	if (!reflection || reflection.type === 'none') return ''

	let blurRad = valToPts(reflection.blur || 0)
	let dist = valToPts(reflection.distance || 0)
	let stA = Math.round((100 - (reflection.transparency || 0)) * 1000)
	let endA = 0
	let sz = Math.round((reflection.size || 100) * 1000)

	// Preset reflection types
	switch (reflection.type) {
		case 'tight':
			sz = 40000
			dist = 0
			break
		case 'half':
			sz = 50000
			dist = 0
			break
		case 'full':
			sz = 100000
			dist = 0
			break
	}

	return `<a:reflection blurRad="${blurRad}" stA="${stA}" endA="${endA}" endPos="${sz}" dist="${dist}" dir="5400000" fadeDir="5400000" algn="bl" rotWithShape="0"/>`
}

/**
 * Generate glow effect XML
 * @param {import('./core-interfaces').GlowProps} glow - glow properties
 * @returns {string} XML string
 */
export function genXmlGlowEffect (glow: any): string {
	if (!glow || !glow.size || !glow.color) return ''

	const rad = valToPts(glow.size)
	let innerElements = ''
	if (glow.transparency) {
		innerElements += `<a:alpha val="${Math.round((100 - glow.transparency) * 1000)}"/>`
	}

	return `<a:glow rad="${rad}">${createColorElement(glow.color, innerElements)}</a:glow>`
}

/**
 * Generate soft edge effect XML
 * @param {import('./core-interfaces').SoftEdgeProps} softEdge - soft edge properties
 * @returns {string} XML string
 */
export function genXmlSoftEdge (softEdge: any): string {
	if (!softEdge || !softEdge.size) return ''
	const rad = valToPts(softEdge.size)
	return `<a:softEdge rad="${rad}"/>`
}

/**
 * Generate 3D shape properties XML
 * @param {import('./core-interfaces').Shape3DProps} shape3D - 3D properties
 * @returns {string} XML string
 */
export function genXmlShape3D (shape3D: any): string {
	if (!shape3D) return ''

	let xml = '<a:sp3d'
	if (shape3D.extrusionHeight) xml += ` extrusionH="${valToPts(shape3D.extrusionHeight)}"`
	if (shape3D.contourWidth) xml += ` contourW="${valToPts(shape3D.contourWidth)}"`
	if (shape3D.material) xml += ` prstMaterial="${shape3D.material}"`
	xml += '>'

	// Top bevel
	if (shape3D.topBevel) {
		xml += `<a:bevelT w="${valToPts(shape3D.topBevel.width || 6)}" h="${valToPts(shape3D.topBevel.height || 6)}" prst="${shape3D.topBevel.type}"/>`
	}

	// Bottom bevel
	if (shape3D.bottomBevel) {
		xml += `<a:bevelB w="${valToPts(shape3D.bottomBevel.width || 6)}" h="${valToPts(shape3D.bottomBevel.height || 6)}" prst="${shape3D.bottomBevel.type}"/>`
	}

	// Extrusion color
	if (shape3D.extrusionColor) {
		xml += '<a:extrusionClr>'
		xml += createColorElement(shape3D.extrusionColor)
		xml += '</a:extrusionClr>'
	}

	// Contour color
	if (shape3D.contourColor) {
		xml += '<a:contourClr>'
		xml += createColorElement(shape3D.contourColor)
		xml += '</a:contourClr>'
	}

	xml += '</a:sp3d>'

	// Camera and lighting - always add scene3d for 3D effects to work properly
	xml += '<a:scene3d><a:camera prst="orthographicFront"/>'

	if (shape3D.lighting) {
		xml += `<a:lightRig rig="${shape3D.lighting}" dir="t"`
		if (shape3D.lightingAngle) {
			xml += `><a:rot lat="${convertRotationDegrees(shape3D.lightingAngle)}" lon="0" rev="0"/></a:lightRig>`
		} else {
			xml += '/>'
		}
	} else {
		// Default lighting if not specified
		xml += '<a:lightRig rig="threePt" dir="t"/>'
	}

	xml += '</a:scene3d>'

	return xml
}

/**
 * Create color selection (ENHANCED VERSION with gradient, pattern, texture support)
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @param {number} textureRId - optional relationship ID for texture fills
 * @returns XML string
 */
export function genXmlColorSelection (props: Color | ShapeFillProps | ShapeLineProps, textureRId?: number): string {
	let fillType = 'solid'
	let colorVal = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		if (typeof props === 'string') colorVal = props
		else {
			if (props.type) fillType = props.type
			if (props.color) colorVal = props.color
			if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
			if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1000)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			case 'gradient':
				if (typeof props !== 'string' && (props as any).gradient) {
					outText += genXmlGradientFill((props as any).gradient)
				}
				break
			case 'pattern':
				if (typeof props !== 'string' && (props as any).pattern) {
					outText += genXmlPatternFill((props as any).pattern)
				}
				break
			case 'texture':
				if (typeof props !== 'string' && (props as any).texture && textureRId) {
					outText += genXmlTextureFill((props as any).texture, textureRId)
				}
				break
			default: // @note need a statement as having only "break" is removed by rollup, then tiggers "no-default" js-linter
				outText += ''
				break
		}
	}

	return outText
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId (target: PresSlide): number {
	// NOTE: defaultRels in slideObjectRelationsToXml always includes slideLayout(rId1) + notesSlide(rId2), hence +2
	return target._rels.length + target._relsChart.length + target._relsMedia.length + 2 + 1
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions (ShadowProps: ShadowProps): ShadowProps | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		// console.warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return
	}

	// OPT: `type`
	if (ShadowProps.type !== 'outer' && ShadowProps.type !== 'inner' && ShadowProps.type !== 'none') {
		console.warn('Warning: shadow.type options are `outer`, `inner` or `none`.')
		ShadowProps.type = 'outer'
	}

	// OPT: `angle`
	if (ShadowProps.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.angle)) || ShadowProps.angle < 0 || ShadowProps.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			ShadowProps.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.angle = Math.round(Number(ShadowProps.angle))
	}

	// OPT: `opacity`
	if (ShadowProps.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.opacity)) || ShadowProps.opacity < 0 || ShadowProps.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			ShadowProps.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.opacity = Number(ShadowProps.opacity)
	}

	// OPT: `color`
	if (ShadowProps.color) {
		// INCORRECT FORMAT
		if (ShadowProps.color.startsWith('#')) {
			console.warn('Warning: shadow.color should not include hash (#) character, , e.g. "FF0000"')
			ShadowProps.color = ShadowProps.color.replace('#', '')
		}
	}

	return ShadowProps
}
