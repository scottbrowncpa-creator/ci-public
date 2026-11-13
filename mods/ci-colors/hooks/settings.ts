// Settings for the ci-colors mod. Edit a value, save, and the next reply draws with it.
// Font face and font size are not here: no mod can set them. They belong to the app you read in
// (Claude desktop app: Settings > Appearance; Windows Terminal: Settings > your profile > Appearance > Font face / Font size).
import type { Theme } from './theme'

// ---------- Width ----------
export const MAX_WIDTH = 120 // widest a reply, table or code block draws, in columns; 0 = use the full window

// ---------- Text styles (what a mod can control) ----------
export const HEADING_STYLE: 'banner' | 'bold' | 'underline' | 'uppercase' = 'banner' // banner = box around H1, heavy rule under H2
export const BOLD_TEXT = true        // **bold** draws heavy; false = color only
export const ITALIC_TEXT = true      // *italic* draws slanted; false = color only
export const UNDERLINE_LINKS = true  // links draw underlined; false = color only

// ---------- Colors ----------
// Palette: teal #1C424D, green #7DB03D, gold #F5C449, charcoal #444755.
// Readable on a dark screen needs about 4.5x contrast. Teal (1.5x) and charcoal (1.8x) fail as text,
// so text uses a brighter teal of the same hue; charcoal is kept only for faint divider lines.
// Delete a line to fall back to the base theme (catppuccin-mocha) for that item.
export const USE_BRAND = true // false = ignore the colors below and use the plain base theme

export const BRAND: Theme = {
  heading: '#5BB0C4',      // H1 and H2 headings - brighter brand teal (6.6x on dark)
  accent: '#7DB03D',       // H3+ headings, quote bars - brand green (6.4x)
  bullet: '#7DB03D',       // list bullets and numbers - brand green
  strong: '#F5C449',       // **bold** text - brand gold (10.1x)
  tableHeader: '#F5C449',  // table header cells - brand gold
  tableRule: '#3D8FA3',    // table lines - deeper brand teal, quieter than headings
  rule: '#444755',         // horizontal rules and chart gridlines - brand charcoal (faint on purpose)
  number: '#7DB03D',       // dollar amounts, counts, versions - brand green
  link: '#5BB0C4',         // links and URLs - brighter brand teal
  path: '#5BB0C4',         // file paths - brighter brand teal
  inlineCode: '#5BB0C4',   // `inline code` - brighter brand teal
  diagram: '#5BB0C4',      // mermaid boxes and arrows - brighter brand teal
  promptBorder: '#7DB03D', // frame around your own prompts - brand green
  replyBorder: '#F5C449',  // frame and label on Claude's replies - brand gold
}
