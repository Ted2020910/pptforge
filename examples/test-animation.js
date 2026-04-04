/**
 * PptxGenJS Animation Test
 * Tests all animation features: entrance, exit, transitions, byParagraph, trigger modes
 */

import pptxgen from '../src/bld/pptxgen.es.js'

async function main() {
	const pptx = new pptxgen()
	pptx.layout = 'LAYOUT_16x9'

	// ========== Slide 1: P0 entrance animations ==========
	const slide1 = pptx.addSlide()
	slide1.background = { color: 'F5F5F5' }
	slide1.transition = { type: 'fade', speed: 'med' }

	slide1.addText('P0 Entrance Animations', {
		x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: '333333',
	})

	// 1. Appear
	slide1.addText('1. Appear', {
		x: 0.5, y: 1.2, w: 4, h: 0.5, fontSize: 16, color: '0066CC',
		animation: { type: 'appear', trigger: 'onClick' },
	})

	// 2. Fade In
	slide1.addText('2. Fade In', {
		x: 0.5, y: 1.9, w: 4, h: 0.5, fontSize: 16, color: '009933',
		animation: { type: 'fadeIn', dur: 800, trigger: 'onClick' },
	})

	// 3. Fly In from bottom
	slide1.addText('3. Fly In (bottom)', {
		x: 0.5, y: 2.6, w: 4, h: 0.5, fontSize: 16, color: 'CC3300',
		animation: { type: 'flyIn', dir: 'bottom', dur: 500, trigger: 'onClick' },
	})

	// 4. Wipe from left
	slide1.addText('4. Wipe (left)', {
		x: 0.5, y: 3.3, w: 4, h: 0.5, fontSize: 16, color: '9900CC',
		animation: { type: 'wipe', dir: 'left', dur: 600, trigger: 'onClick' },
	})

	// 5. Fly In from right
	slide1.addText('5. Fly In (right)', {
		x: 5, y: 1.2, w: 4, h: 0.5, fontSize: 16, color: 'FF6600',
		animation: { type: 'flyIn', dir: 'right', dur: 500, trigger: 'onClick' },
	})

	// 6. Wipe (down)
	slide1.addText('6. Wipe (down)', {
		x: 5, y: 1.9, w: 4, h: 0.5, fontSize: 16, color: '006666',
		animation: { type: 'wipe', dir: 'down', dur: 700, trigger: 'onClick' },
	})

	// ========== Slide 2: P1 filter animations ==========
	const slide2 = pptx.addSlide()
	slide2.background = { color: 'FFFFFF' }
	slide2.transition = { type: 'wipe', dir: 'left', speed: 'fast' }

	slide2.addText('P1 Filter Animations', {
		x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: '333333',
	})

	// Blinds
	slide2.addText('Blinds (horizontal)', {
		x: 0.5, y: 1.2, w: 4, h: 0.5, fontSize: 14, color: '0066CC',
		animation: { type: 'blinds', dir: 'horizontal', dur: 500, trigger: 'onClick' },
	})

	// Box
	slide2.addText('Box (in)', {
		x: 0.5, y: 1.9, w: 4, h: 0.5, fontSize: 14, color: '009933',
		animation: { type: 'box', inOut: 'in', dur: 1000, trigger: 'onClick' },
	})

	// Diamond
	slide2.addText('Diamond (in)', {
		x: 0.5, y: 2.6, w: 4, h: 0.5, fontSize: 14, color: 'CC3300',
		animation: { type: 'diamond', inOut: 'in', dur: 1500, trigger: 'onClick' },
	})

	// Circle
	slide2.addText('Circle (in)', {
		x: 5, y: 1.2, w: 4, h: 0.5, fontSize: 14, color: '9900CC',
		animation: { type: 'circle', inOut: 'in', dur: 1200, trigger: 'onClick' },
	})

	// Dissolve
	slide2.addText('Dissolve', {
		x: 5, y: 1.9, w: 4, h: 0.5, fontSize: 14, color: 'FF6600',
		animation: { type: 'dissolve', dur: 800, trigger: 'onClick' },
	})

	// Wedge
	slide2.addText('Wedge', {
		x: 5, y: 2.6, w: 4, h: 0.5, fontSize: 14, color: '006666',
		animation: { type: 'wedge', dur: 1000, trigger: 'onClick' },
	})

	// ========== Slide 3: withPrev & afterPrev ==========
	const slide3 = pptx.addSlide()
	slide3.background = { color: 'EEEEFF' }
	slide3.transition = { type: 'push', dir: 'left' }

	slide3.addText('Trigger Modes: withPrev & afterPrev', {
		x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: '333333',
	})

	// These 3 elements animate together on first click
	slide3.addText('Click once -> all 3 appear simultaneously', {
		x: 0.5, y: 1.0, w: 9, h: 0.4, fontSize: 12, italic: true, color: '666666',
	})

	slide3.addText('Element A (onClick)', {
		x: 0.5, y: 1.5, w: 2.5, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '0066CC' },
		animation: { type: 'fadeIn', dur: 500, trigger: 'onClick', order: 1 },
	})

	slide3.addText('Element B (withPrev)', {
		x: 3.5, y: 1.5, w: 2.5, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '009933' },
		animation: { type: 'fadeIn', dur: 500, trigger: 'withPrev', order: 2 },
	})

	slide3.addText('Element C (withPrev)', {
		x: 6.5, y: 1.5, w: 2.5, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: 'CC3300' },
		animation: { type: 'fadeIn', dur: 500, trigger: 'withPrev', order: 3 },
	})

	// Staggered animation with afterPrev
	slide3.addText('Click once -> staggered sequence:', {
		x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 12, italic: true, color: '666666',
	})

	slide3.addText('Step 1', {
		x: 0.5, y: 3.0, w: 2, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '336699' },
		animation: { type: 'flyIn', dir: 'bottom', dur: 400, trigger: 'onClick', order: 4 },
	})

	slide3.addText('Step 2', {
		x: 2.8, y: 3.0, w: 2, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '339966' },
		animation: { type: 'flyIn', dir: 'bottom', dur: 400, trigger: 'afterPrev', delay: 100, order: 5 },
	})

	slide3.addText('Step 3', {
		x: 5.1, y: 3.0, w: 2, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '996633' },
		animation: { type: 'flyIn', dir: 'bottom', dur: 400, trigger: 'afterPrev', delay: 100, order: 6 },
	})

	slide3.addText('Step 4', {
		x: 7.4, y: 3.0, w: 2, h: 0.5, fontSize: 14, color: 'FFFFFF',
		fill: { color: '993366' },
		animation: { type: 'flyIn', dir: 'bottom', dur: 400, trigger: 'afterPrev', delay: 100, order: 7 },
	})

	// ========== Slide 4: byParagraph ==========
	const slide4 = pptx.addSlide()
	slide4.background = { color: 'FFFFF0' }

	slide4.addText('byParagraph: Text appears line by line', {
		x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: '333333',
	})

	slide4.addText([
		{ text: 'Point 1: First important item', options: { fontSize: 16, bullet: true } },
		{ text: 'Point 2: Second important item', options: { fontSize: 16, bullet: true } },
		{ text: 'Point 3: Third important item', options: { fontSize: 16, bullet: true } },
		{ text: 'Point 4: Fourth important item', options: { fontSize: 16, bullet: true } },
	], {
		x: 0.5, y: 1.5, w: 8, h: 3,
		color: '333333',
		animation: { type: 'wipe', dir: 'left', dur: 500, trigger: 'onClick', byParagraph: true },
	})

	// ========== Slide 5: Image with animation ==========
	const slide5 = pptx.addSlide()
	slide5.background = { color: 'F0F0FF' }

	slide5.addText('Images with Animation', {
		x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: '333333',
	})

	slide5.addShape(pptx.shapes.RECTANGLE, {
		x: 1, y: 1.5, w: 3, h: 2.5,
		fill: { color: '0066CC' },
		animation: { type: 'flyIn', dir: 'left', dur: 600, trigger: 'onClick' },
	})

	slide5.addShape(pptx.shapes.OVAL, {
		x: 5, y: 1.5, w: 3, h: 2.5,
		fill: { color: 'CC3300' },
		animation: { type: 'diamond', inOut: 'in', dur: 1000, trigger: 'onClick' },
	})

	// ========== Save ==========
	const filename = 'test-animation.pptx'
	await pptx.writeFile({ fileName: `examples/${filename}` })
	console.log(`Created: examples/${filename}`)
	console.log('Open in PowerPoint to verify animations!')
}

main().catch(console.error)
