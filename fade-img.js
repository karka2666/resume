const sharp = require('sharp');

async function fade() {
  const bg = await sharp('D:/portfolio/images/hero-bg.png').toBuffer();
  const meta = await sharp(bg).metadata();

  // Create white overlay with 78% opacity (lighter background)
  const overlay = await sharp({
    create: { width: meta.width, height: meta.height, channels: 4, background: { r: 253, g: 249, b: 248, alpha: 0.78 } }
  }).png().toBuffer();

  await sharp(bg)
    .composite([{ input: overlay, blend: 'over' }])
    .png()
    .toFile('D:/portfolio/images/hero-bg-faded.png');

  console.log('Faded image created: ' + meta.width + 'x' + meta.height);
}

fade().catch(e => console.error(e.message));
