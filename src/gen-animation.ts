/**
 * PptxGenJS: Animation XML Generation
 * Generates <p:timing> and <p:transition> XML for slide animations
 */

import { AnimationProps, AnimationType, ISlideObject, PresSlide, TransitionProps } from './core-interfaces'
import { SLIDE_OBJECT_TYPES } from './core-enums'

// =============================== TYPES ===============================

interface AnimationEntry {
	spid: number
	anim: AnimationProps
	isTextElement: boolean
	paragraphRange?: { st: number; end: number }
}

interface PresetConfig {
	presetID: number
	defaultClass: 'entr' | 'exit'
	method: 'set' | 'animEffect' | 'anim'
	filter?: (dir?: string, inOut?: string) => string
	defaultSubtype: number
	subtypeMap?: Record<string, number>
}

// =============================== PRESET MAPPINGS ===============================

const ANIMATION_PRESETS: Record<string, PresetConfig> = {
	// P0 entrance
	appear: { presetID: 1, defaultClass: 'entr', method: 'set', defaultSubtype: 0 },
	fadeIn: { presetID: 10, defaultClass: 'entr', method: 'animEffect', filter: () => 'fade', defaultSubtype: 0 },
	flyIn: {
		presetID: 2,
		defaultClass: 'entr',
		method: 'anim',
		defaultSubtype: 4,
		subtypeMap: { bottom: 4, top: 2, left: 8, right: 6 },
	},
	wipe: {
		presetID: 22,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `wipe(${dir || 'down'})`,
		defaultSubtype: 4,
		subtypeMap: { down: 4, left: 2, right: 8, up: 12 },
	},
	zoom: { presetID: 53, defaultClass: 'entr', method: 'animEffect', filter: () => 'fade', defaultSubtype: 0 },

	// P0 exit
	disappear: { presetID: 1, defaultClass: 'exit', method: 'set', defaultSubtype: 0 },
	fadeOut: { presetID: 10, defaultClass: 'exit', method: 'animEffect', filter: () => 'fade', defaultSubtype: 0 },
	flyOut: {
		presetID: 2,
		defaultClass: 'exit',
		method: 'anim',
		defaultSubtype: 4,
		subtypeMap: { bottom: 4, top: 2, left: 8, right: 6 },
	},

	// P1 filter extensions
	blinds: {
		presetID: 3,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `blinds(${dir || 'horizontal'})`,
		defaultSubtype: 10,
		subtypeMap: { horizontal: 10, vertical: 5 },
	},
	box: {
		presetID: 4,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (_, inOut) => `box(${inOut || 'in'})`,
		defaultSubtype: 16,
	},
	checkerboard: {
		presetID: 5,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `checkerboard(${dir || 'across'})`,
		defaultSubtype: 10,
	},
	circle: {
		presetID: 6,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (_, inOut) => `circle(${inOut || 'in'})`,
		defaultSubtype: 16,
	},
	diamond: {
		presetID: 8,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (_, inOut) => `diamond(${inOut || 'in'})`,
		defaultSubtype: 16,
	},
	dissolve: { presetID: 9, defaultClass: 'entr', method: 'animEffect', filter: () => 'dissolve', defaultSubtype: 0 },
	plus: {
		presetID: 14,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (_, inOut) => `plus(${inOut || 'in'})`,
		defaultSubtype: 16,
	},
	randomBars: {
		presetID: 16,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `randombar(${dir || 'horizontal'})`,
		defaultSubtype: 10,
	},
	split: {
		presetID: 18,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir, inOut) => `split(${dir || 'horizontal'} ${inOut || 'in'})`,
		defaultSubtype: 16,
	},
	strips: {
		presetID: 19,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `strips(${dir || 'downLeft'})`,
		defaultSubtype: 0,
	},
	wedge: { presetID: 21, defaultClass: 'entr', method: 'animEffect', filter: () => 'wedge', defaultSubtype: 0 },
	wheel: {
		presetID: 21,
		defaultClass: 'entr',
		method: 'animEffect',
		filter: (dir) => `wheel(${dir || '1'})`,
		defaultSubtype: 1,
	},
	random: { presetID: 25, defaultClass: 'entr', method: 'animEffect', filter: () => 'random', defaultSubtype: 0 },
}

// FlyIn/FlyOut direction to keyframe values
const FLY_KEYFRAMES: Record<string, { axis: string; from: string; to: string }> = {
	bottom: { axis: 'ppt_y', from: '1+#ppt_h/2', to: '#ppt_y' },
	top: { axis: 'ppt_y', from: '0-#ppt_h', to: '#ppt_y' },
	left: { axis: 'ppt_x', from: '0-#ppt_w', to: '#ppt_x' },
	right: { axis: 'ppt_x', from: '1+#ppt_w/2', to: '#ppt_x' },
}

// =============================== ID COUNTER ===============================

class CtnIdCounter {
	private _id = 0
	next(): number {
		return ++this._id
	}
}

// =============================== HELPER: spTgt XML ===============================

function genSpTgt(spid: number, paragraphRange?: { st: number; end: number }): string {
	if (paragraphRange) {
		return `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${paragraphRange.st}" end="${paragraphRange.end}"/></p:txEl></p:spTgt>`
	}
	return `<p:spTgt spid="${spid}"/>`
}

// =============================== ANIMATION CONTENT GENERATORS ===============================

/**
 * Generate visibility set XML (used by all animations as the first step)
 */
function genVisibilitySet(counter: CtnIdCounter, spid: number, isExit: boolean, paragraphRange?: { st: number; end: number }): string {
	const id = counter.next()
	const val = isExit ? 'hidden' : 'visible'
	return (
		'<p:set>' +
		'<p:cBhvr>' +
		`<p:cTn id="${id}" dur="1" fill="hold">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'</p:cTn>' +
		`<p:tgtEl>${genSpTgt(spid, paragraphRange)}</p:tgtEl>` +
		'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>' +
		'</p:cBhvr>' +
		`<p:to><p:strVal val="${val}"/></p:to>` +
		'</p:set>'
	)
}

/**
 * Generate animEffect XML (for filter-based animations: wipe, blinds, box, etc.)
 */
function genAnimEffect(
	counter: CtnIdCounter,
	spid: number,
	filter: string,
	dur: number,
	isExit: boolean,
	paragraphRange?: { st: number; end: number }
): string {
	const id = counter.next()
	const transition = isExit ? 'out' : 'in'
	return (
		`<p:animEffect transition="${transition}" filter="${filter}">` +
		'<p:cBhvr>' +
		`<p:cTn id="${id}" dur="${dur}"/>` +
		`<p:tgtEl>${genSpTgt(spid, paragraphRange)}</p:tgtEl>` +
		'</p:cBhvr>' +
		'</p:animEffect>'
	)
}

/**
 * Generate fly animation XML (keyframe-based position animation)
 */
function genFlyAnim(
	counter: CtnIdCounter,
	spid: number,
	dir: string,
	dur: number,
	isExit: boolean,
	paragraphRange?: { st: number; end: number }
): string {
	const keyframes = FLY_KEYFRAMES[dir] || FLY_KEYFRAMES.bottom
	const otherAxis = keyframes.axis === 'ppt_x' ? 'ppt_y' : 'ppt_x'

	// For exit animations, reverse from and to
	const fromVal = isExit ? keyframes.to : keyframes.from
	const toVal = isExit ? keyframes.from : keyframes.to
	const otherRef = `#${otherAxis}`

	let xml = ''

	// Axis that stays (no movement)
	const idStay = counter.next()
	xml +=
		'<p:anim calcmode="lin" valueType="num">' +
		'<p:cBhvr additive="base">' +
		`<p:cTn id="${idStay}" dur="${dur}" fill="hold"/>` +
		`<p:tgtEl>${genSpTgt(spid, paragraphRange)}</p:tgtEl>` +
		`<p:attrNameLst><p:attrName>${otherAxis}</p:attrName></p:attrNameLst>` +
		'</p:cBhvr>' +
		'<p:tavLst>' +
		`<p:tav tm="0"><p:val><p:strVal val="${otherRef}"/></p:val></p:tav>` +
		`<p:tav tm="100000"><p:val><p:strVal val="${otherRef}"/></p:val></p:tav>` +
		'</p:tavLst>' +
		'</p:anim>'

	// Axis that moves
	const idMove = counter.next()
	xml +=
		'<p:anim calcmode="lin" valueType="num">' +
		'<p:cBhvr additive="base">' +
		`<p:cTn id="${idMove}" dur="${dur}" fill="hold"/>` +
		`<p:tgtEl>${genSpTgt(spid, paragraphRange)}</p:tgtEl>` +
		`<p:attrNameLst><p:attrName>${keyframes.axis}</p:attrName></p:attrNameLst>` +
		'</p:cBhvr>' +
		'<p:tavLst>' +
		`<p:tav tm="0"><p:val><p:strVal val="${fromVal}"/></p:val></p:tav>` +
		`<p:tav tm="100000"><p:val><p:strVal val="${toVal}"/></p:val></p:tav>` +
		'</p:tavLst>' +
		'</p:anim>'

	return xml
}

// =============================== SINGLE ANIMATION XML ===============================

/**
 * Generate XML for a single animation entry
 */
function genSingleAnimation(counter: CtnIdCounter, entry: AnimationEntry, nodeType: string): string {
	const anim = entry.anim
	const preset = ANIMATION_PRESETS[anim.type]
	if (!preset) return ''

	const presetClass = anim.class === 'exit' ? 'exit' : preset.defaultClass
	const isExit = presetClass === 'exit'
	const dur = anim.dur || 500
	const delay = anim.delay || 0

	// Determine subtype
	let subtype = preset.defaultSubtype
	if (preset.subtypeMap) {
		const key = anim.dir || anim.inOut || ''
		if (key && preset.subtypeMap[key] !== undefined) {
			subtype = preset.subtypeMap[key]
		}
	}

	// Build the inner animation content
	let innerXml = ''

	// Step 1: Visibility set (always present)
	innerXml += genVisibilitySet(counter, entry.spid, isExit, entry.paragraphRange)

	// Step 2: Animation-specific content
	if (preset.method === 'animEffect' && preset.filter) {
		const filterStr = preset.filter(anim.dir, anim.inOut)
		innerXml += genAnimEffect(counter, entry.spid, filterStr, dur, isExit, entry.paragraphRange)
	} else if (preset.method === 'anim') {
		const dir = anim.dir || 'bottom'
		innerXml += genFlyAnim(counter, entry.spid, dir, dur, isExit, entry.paragraphRange)
	}
	// For 'set' method, only the visibility set is needed (already added)

	// Build the animation cTn node
	const animCtnId = counter.next()
	return (
		'<p:par>' +
		`<p:cTn id="${animCtnId}" presetID="${preset.presetID}" presetClass="${presetClass}" presetSubtype="${subtype}" fill="hold" nodeType="${nodeType}">` +
		`<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst>` +
		`<p:childTnLst>${innerXml}</p:childTnLst>` +
		'</p:cTn>' +
		'</p:par>'
	)
}

// =============================== CLICK GROUP LOGIC ===============================

interface ClickGroup {
	entries: { entry: AnimationEntry; nodeType: string }[]
}

/**
 * Group animation entries into click groups based on trigger type
 */
function buildClickGroups(entries: AnimationEntry[]): ClickGroup[] {
	const groups: ClickGroup[] = []
	let currentGroup: ClickGroup | null = null

	for (const entry of entries) {
		const trigger = entry.anim.trigger || 'onClick'

		if (trigger === 'onClick') {
			// Start a new click group
			currentGroup = { entries: [{ entry, nodeType: 'clickEffect' }] }
			groups.push(currentGroup)
		} else if (trigger === 'withPrev') {
			// Add to current group (or create new if none exists)
			if (!currentGroup) {
				currentGroup = { entries: [] }
				groups.push(currentGroup)
			}
			currentGroup.entries.push({ entry, nodeType: 'withEffect' })
		} else if (trigger === 'afterPrev') {
			// Add to current group (or create new if none exists)
			if (!currentGroup) {
				currentGroup = { entries: [] }
				groups.push(currentGroup)
			}
			currentGroup.entries.push({ entry, nodeType: 'afterEffect' })
		}
	}

	return groups
}

/**
 * Generate XML for a single click group
 */
function genClickGroupXml(counter: CtnIdCounter, group: ClickGroup): string {
	// Generate all animation XMLs in this group
	let animationsXml = ''
	for (const { entry, nodeType } of group.entries) {
		animationsXml += genSingleAnimation(counter, entry, nodeType)
	}

	const outerCtnId = counter.next()
	const innerCtnId = counter.next()

	return (
		'<p:par>' +
		`<p:cTn id="${outerCtnId}" fill="hold">` +
		'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${innerCtnId}" fill="hold">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		`<p:childTnLst>${animationsXml}</p:childTnLst>` +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>'
	)
}

// =============================== MAIN FUNCTIONS ===============================

/**
 * Collect animation entries from slide objects
 */
function collectAnimationEntries(slide: PresSlide): AnimationEntry[] {
	const entries: AnimationEntry[] = []

	if (!slide._slideObjects) return entries

	for (const obj of slide._slideObjects) {
		if (!obj.options?.animation || !obj._animShapeId) continue

		const animations = Array.isArray(obj.options.animation) ? obj.options.animation : [obj.options.animation]

		const isTextElement = obj._type === SLIDE_OBJECT_TYPES.text || obj._type === SLIDE_OBJECT_TYPES.formula

		for (const anim of animations) {
			if (!anim || !anim.type) continue

			if (anim.byParagraph && isTextElement && obj.text && obj.text.length > 1) {
				// Expand into per-paragraph animations
				for (let i = 0; i < obj.text.length; i++) {
					entries.push({
						spid: obj._animShapeId,
						anim: {
							...anim,
							trigger: i === 0 ? (anim.trigger || 'onClick') : 'onClick',
						},
						isTextElement: true,
						paragraphRange: { st: i, end: i },
					})
				}
			} else {
				entries.push({
					spid: obj._animShapeId,
					anim,
					isTextElement,
				})
			}
		}
	}

	// Sort by order (if specified), preserving original order for unspecified
	entries.sort((a, b) => {
		const orderA = a.anim.order ?? Infinity
		const orderB = b.anim.order ?? Infinity
		return orderA - orderB
	})

	return entries
}

// =============================== MEDIA ENTRIES ===============================

interface MediaEntry {
	spid: number
	autoPlay: boolean
	loop: boolean
}

/**
 * Collect media entries that need timing XML (autoPlay/loop videos)
 */
function collectMediaEntries(slide: PresSlide): MediaEntry[] {
	const entries: MediaEntry[] = []

	if (!slide._slideObjects) return entries

	for (const obj of slide._slideObjects) {
		if (obj._type !== SLIDE_OBJECT_TYPES.media) continue
		if (obj.mtype !== 'video') continue
		if (!obj._animShapeId) continue

		const autoPlay = obj.options?.autoPlay === true
		const loop = obj.options?.loop === true

		if (autoPlay || loop) {
			entries.push({
				spid: obj._animShapeId,
				autoPlay,
				loop,
			})
		}
	}

	return entries
}

// =============================== MEDIA TIMING GENERATORS ===============================

/**
 * Generate the playFrom(0.0) click group for autoplay video.
 * Uses dual stCondLst: delay="indefinite" + evt="onBegin" referencing mainSeq
 * to trigger automatically when slide begins.
 */
function genMediaPlayCommand(counter: CtnIdCounter, spid: number, mainSeqId: number): string {
	const outerCtnId = counter.next()
	const innerCtnId = counter.next()
	const presetCtnId = counter.next()
	const cmdCtnId = counter.next()

	return (
		'<p:par>' +
		`<p:cTn id="${outerCtnId}" fill="hold">` +
		'<p:stCondLst>' +
		'<p:cond delay="indefinite"/>' +
		`<p:cond evt="onBegin" delay="0"><p:tn val="${mainSeqId}"/></p:cond>` +
		'</p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${innerCtnId}" fill="hold">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${presetCtnId}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:cmd type="call" cmd="playFrom(0.0)">' +
		'<p:cBhvr additive="base">' +
		`<p:cTn id="${cmdCtnId}" dur="1" fill="hold"/>` +
		`<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
		'</p:cBhvr>' +
		'</p:cmd>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>'
	)
}

/**
 * Generate <p:video> node for loop control.
 * Placed as direct child of tmRoot (sibling of mainSeq).
 */
function genMediaVideoNode(counter: CtnIdCounter, spid: number, loop: boolean): string {
	const ctnId = counter.next()
	const repeatAttr = loop ? ' repeatCount="indefinite"' : ''

	return (
		'<p:video fullScrn="0">' +
		'<p:cMediaNode>' +
		`<p:cTn id="${ctnId}"${repeatAttr} fill="hold" display="1">` +
		'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>' +
		'</p:cTn>' +
		`<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
		'</p:cMediaNode>' +
		'</p:video>'
	)
}

/**
 * Generate interactiveSeq for click-to-toggle-pause on video.
 * Placed as direct child of tmRoot (sibling of mainSeq).
 */
function genMediaInteractiveSeq(counter: CtnIdCounter, spid: number): string {
	const seqCtnId = counter.next()
	const outerCtnId = counter.next()
	const innerCtnId = counter.next()
	const presetCtnId = counter.next()
	const cmdCtnId = counter.next()

	return (
		'<p:seq concurrent="1" nextAc="seek">' +
		`<p:cTn id="${seqCtnId}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">` +
		`<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:stCondLst>` +
		'<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${outerCtnId}" fill="hold">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${innerCtnId}" fill="hold">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:par>' +
		`<p:cTn id="${presetCtnId}" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect">` +
		'<p:stCondLst><p:cond delay="0"/></p:stCondLst>' +
		'<p:childTnLst>' +
		'<p:cmd type="call" cmd="togglePause">' +
		'<p:cBhvr additive="base">' +
		`<p:cTn id="${cmdCtnId}" dur="1" fill="hold"/>` +
		`<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
		'</p:cBhvr>' +
		'</p:cmd>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		'</p:par>' +
		'</p:childTnLst>' +
		'</p:cTn>' +
		`<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cond></p:nextCondLst>` +
		'</p:seq>'
	)
}

/**
 * Generate <p:timing> XML for a slide
 * Supports both element animations and video media autoplay/loop timing.
 * When both exist, they are merged into a single <p:timing> root.
 * @param {PresSlide} slide - the slide containing animated objects
 * @returns {string} timing XML or empty string
 */
export function makeXmlTiming(slide: PresSlide): string {
	const animEntries = collectAnimationEntries(slide)
	const mediaEntries = collectMediaEntries(slide)

	const hasAnimations = animEntries.length > 0
	const hasMedia = mediaEntries.length > 0

	if (!hasAnimations && !hasMedia) return ''

	const counter = new CtnIdCounter()

	// Root timing node
	const rootId = counter.next() // id=1
	const mainSeqId = counter.next() // id=2

	// ---- Build mainSeq childTnLst (animation click groups + media play commands) ----
	let mainSeqChildXml = ''

	// 1. Animation click groups
	if (hasAnimations) {
		const groups = buildClickGroups(animEntries)
		for (const group of groups) {
			mainSeqChildXml += genClickGroupXml(counter, group)
		}
	}

	// 2. Media autoplay commands (inside mainSeq, triggered by onBegin)
	if (hasMedia) {
		for (const media of mediaEntries) {
			if (media.autoPlay) {
				mainSeqChildXml += genMediaPlayCommand(counter, media.spid, mainSeqId)
			}
		}
	}

	// If mainSeq has no children, we still need to emit timing for media loop/interactive
	// Use an empty childTnLst in that case
	const mainSeqXml = mainSeqChildXml
		? `<p:childTnLst>${mainSeqChildXml}</p:childTnLst>`
		: '<p:childTnLst/>'

	// ---- Build tmRoot children (mainSeq + media video nodes + interactive seqs) ----
	let tmRootChildXml =
		'<p:seq concurrent="1" nextAc="seek">' +
		`<p:cTn id="${mainSeqId}" dur="indefinite" nodeType="mainSeq">` +
		mainSeqXml +
		'</p:cTn>' +
		'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
		'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>' +
		'</p:seq>'

	// 3. Media video nodes (loop control) and interactive seqs (click-to-pause)
	if (hasMedia) {
		for (const media of mediaEntries) {
			tmRootChildXml += genMediaVideoNode(counter, media.spid, media.loop)
			tmRootChildXml += genMediaInteractiveSeq(counter, media.spid)
		}
	}

	return (
		'<p:timing>' +
		'<p:tnLst>' +
		'<p:par>' +
		`<p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot">` +
		`<p:childTnLst>${tmRootChildXml}</p:childTnLst>` +
		'</p:cTn>' +
		'</p:par>' +
		'</p:tnLst>' +
		'</p:timing>'
	)
}

// =============================== TRANSITION ===============================

const TRANSITION_DIR_MAP: Record<string, string> = {
	left: 'l',
	right: 'r',
	top: 'u',
	bottom: 'd',
}

/**
 * Generate <p:transition> XML for a slide
 * @param {TransitionProps} transition - transition configuration
 * @returns {string} transition XML or empty string
 */
export function makeXmlTransition(transition?: TransitionProps): string {
	if (!transition) return ''
	if (transition.type === 'none') return ''

	// Speed attribute
	const speed = transition.speed || 'med'

	// Build attributes
	let attrs = ` spd="${speed}"`

	// advClick
	if (transition.advOnClick === false) {
		attrs += ' advClick="0"'
	}

	// advTm (auto-advance)
	if (transition.advAfter && transition.advAfter > 0) {
		attrs += ` advTm="${transition.advAfter}"`
	}

	// Duration override (if dur is specified, convert to speed equivalent)
	if (transition.dur) {
		// Remove speed attr and use dur instead
		attrs = ''
		if (transition.advOnClick === false) attrs += ' advClick="0"'
		if (transition.advAfter && transition.advAfter > 0) attrs += ` advTm="${transition.advAfter}"`
	}

	// Build inner element based on type
	let innerXml = ''
	const dir = transition.dir ? TRANSITION_DIR_MAP[transition.dir] || '' : ''

	switch (transition.type || 'fade') {
		case 'fade':
			innerXml = '<p:fade/>'
			break
		case 'push':
			innerXml = `<p:push${dir ? ` dir="${dir}"` : ''}/>`
			break
		case 'wipe':
			innerXml = `<p:wipe${dir ? ` dir="${dir}"` : ''}/>`
			break
		case 'split':
			innerXml = '<p:split/>'
			break
		case 'cut':
			innerXml = '<p:cut/>'
			break
		case 'cover':
			innerXml = `<p:cover${dir ? ` dir="${dir}"` : ''}/>`
			break
		case 'uncover':
			innerXml = `<p:uncover${dir ? ` dir="${dir}"` : ''}/>`
			break
		default:
			innerXml = '<p:fade/>'
	}

	return `<p:transition${attrs}>${innerXml}</p:transition>`
}
