/**
 * PptxGenJS Video Background Test
 * Tests video autoPlay + loop as slide background
 * Uses remote video URL from Pixabay
 */

import pptxgen from '../src/bld/pptxgen.es.js'

async function main() {
	const pptx = new pptxgen()
	pptx.layout = 'LAYOUT_16x9'

	const videoUrl = 'https://cdn.pixabay.com/video/2021/06/06/76681-559745365_medium.mp4?download'

	// ========== Slide 1: Video background with autoPlay + loop ==========
	const slide1 = pptx.addSlide()

	// Add full-screen video as background (first element = bottom layer)
	slide1.addMedia({
		type: 'video',
		path: videoUrl,
		x: 0,
		y: 0,
		w: '100%',
		h: '100%',
		autoPlay: true,
		loop: true,
	})

	// Add text on top of video
	slide1.addText('Video Background Demo', {
		x: 0.5,
		y: 0.5,
		w: 9,
		h: 1,
		fontSize: 36,
		bold: true,
		color: 'FFFFFF',
	})

	slide1.addText('This video auto-plays and loops indefinitely', {
		x: 0.5,
		y: 4.5,
		w: 9,
		h: 0.6,
		fontSize: 18,
		color: 'FFFFFF',
		italic: true,
	})

	// ========== Slide 2: Video with autoPlay only (no loop) ==========
	const slide2 = pptx.addSlide()
	slide2.background = { color: '222222' }

	slide2.addMedia({
		type: 'video',
		path: videoUrl,
		x: 1,
		y: 1,
		w: 8,
		h: 4.5,
		autoPlay: true,
		loop: false,
	})

	slide2.addText('AutoPlay only (no loop)', {
		x: 0.5,
		y: 5.8,
		w: 9,
		h: 0.5,
		fontSize: 16,
		color: 'FFFFFF',
	})

	// ========== Slide 3: Video background + animated text ==========
	const slide3 = pptx.addSlide()

	slide3.addMedia({
		type: 'video',
		path: videoUrl,
		x: 0,
		y: 0,
		w: '100%',
		h: '100%',
		autoPlay: true,
		loop: true,
	})

	slide3.addText('Animation + Video BG', {
		x: 0.5,
		y: 1,
		w: 9,
		h: 1,
		fontSize: 36,
		bold: true,
		color: 'FFFFFF',
		animation: { type: 'fadeIn', dur: 800, trigger: 'onClick' },
	})

	slide3.addText('Elements animate on top of video', {
		x: 0.5,
		y: 2.5,
		w: 9,
		h: 0.6,
		fontSize: 18,
		color: 'FFFF00',
		animation: { type: 'flyIn', dir: 'bottom', dur: 500, trigger: 'afterPrev', delay: 300 },
	})

	// ========== Save ==========
	const filename = 'test-video-bg.pptx'
	await pptx.writeFile({ fileName: `examples/${filename}` })
	console.log(`Created: examples/${filename}`)
	console.log('Open in PowerPoint to verify video auto-plays and loops!')
}

main().catch(console.error)
