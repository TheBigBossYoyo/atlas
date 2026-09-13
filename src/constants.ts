export const SAMPLE_MARKDOWN = `# Welcome to Atlas

A universal document viewer for Word, Excel, PowerPoint, PDF, code, markdown, and more.

---

## Features

### 📂 File Management
- **Open** any \`.md\` file from your computer
- **Drag & drop** files directly into the app
- Instantly renders your markdown content

### 🎨 Themes
Cycle through **5 themes**: Light, Dark, Sepia, Nord, and Dracula. Your preference is saved automatically. Press \`Ctrl+T\` to cycle.

### 📑 Table of Contents
Auto-generated sidebar navigation from your document's headings. Click any heading to jump directly to it.

### 🔍 Search
Press \`Ctrl+F\` (or \`⌘+F\` on Mac) to search within your document. Navigate between matches with arrow buttons.

### 👁️ View Modes
- **Preview** — Clean rendered markdown (default)
- **Editor** — Raw markdown source
- **Split** — Side-by-side editor and preview

### 🖨️ Export
Export to **PDF**, **DOCX**, **HTML**, or **Markdown**. Press \`Ctrl+E\` to open the export menu.

### ✏️ Editing & Autosave
Edit your document live in the editor or split view. Drafts are autosaved to your browser, and \`Ctrl+S\` saves to disk (\`Ctrl+Shift+S\` for Save As).

### 📐 Diagrams (Mermaid)
Render flowcharts, sequence diagrams, and more directly inside your markdown using \`\`\`mermaid code blocks.

### 🔠 Font Size
Adjust reading comfort with \`Ctrl++\` / \`Ctrl+-\` / \`Ctrl+0\`.

### ⌨️ Shortcuts
Press \`Ctrl+/\` to view all keyboard shortcuts.

---

## Markdown Showcase

### Text Formatting

This is **bold text**, this is *italic text*, and this is ~~strikethrough~~.

You can also use \`inline code\` within paragraphs.

### Links & Images

[Visit GitHub](https://github.com) | [Atlas Docs](#)

![Placeholder Image](https://via.placeholder.com/600x200/0969da/ffffff?text=Atlas)

### Blockquotes

> "The best way to predict the future is to invent it."
> — Alan Kay

> **Note:** Blockquotes can contain **formatted text** and even
> - nested
> - lists

### Lists

#### Ordered List
1. First item
2. Second item
3. Third item
   1. Sub-item A
   2. Sub-item B

#### Unordered List
- React + Vite + TypeScript
- Syntax highlighting
- Math/LaTeX support
- Dark/Light themes

#### Task List
- [x] Project setup
- [x] Markdown rendering
- [x] Syntax highlighting
- [x] Theme toggle
- [x] Table of Contents
- [x] Search functionality
- [x] Print/Export PDF
- [x] Math/LaTeX support

### Code Blocks

#### JavaScript
\`\`\`javascript
function fibonacci(n) {
  if (n <= 1) return n;
  return fibonacci(n - 1) + fibonacci(n - 2);
}

// Generate first 10 numbers
const sequence = Array.from({ length: 10 }, (_, i) => fibonacci(i));
console.log(sequence); // [0, 1, 1, 2, 3, 5, 8, 13, 21, 34]
\`\`\`

#### TypeScript
\`\`\`typescript
interface User {
  id: number;
  name: string;
  email: string;
  roles: readonly string[];
}

async function fetchUser(id: number): Promise<User> {
  const response = await fetch(\`/api/users/\${id}\`);
  if (!response.ok) throw new Error('User not found');
  return response.json();
}
\`\`\`

#### Python
\`\`\`python
from dataclasses import dataclass
from typing import List

@dataclass
class Point:
    x: float
    y: float

    def distance_to(self, other: 'Point') -> float:
        return ((self.x - other.x) ** 2 + (self.y - other.y) ** 2) ** 0.5

points: List[Point] = [Point(0, 0), Point(3, 4), Point(1, 1)]
\`\`\`

#### CSS
\`\`\`css
.container {
  display: grid;
  grid-template-columns: 260px 1fr;
  gap: 1rem;
  height: 100vh;
  background: var(--bg-primary);
  color: var(--text-primary);
  transition: background 0.3s ease, color 0.3s ease;
}
\`\`\`

### Tables

| Feature | Status | Notes |
|---------|--------|-------|
| GFM Support | ✅ | Tables, task lists, strikethrough |
| Syntax Highlighting | ✅ | All major languages |
| Math/LaTeX | ✅ | Inline and block equations |
| Dark Mode | ✅ | Persistent preference |
| Search | ✅ | Ctrl+F with match navigation |
| PDF Export | ✅ | Via print dialog |

### Mathematics

#### Inline Math
The quadratic formula is $x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$ and Euler's identity is $e^{i\\pi} + 1 = 0$.

#### Block Math
$$
\\int_{-\\infty}^{\\infty} e^{-x^2} dx = \\sqrt{\\pi}
$$

$$
\\sum_{n=1}^{\\infty} \\frac{1}{n^2} = \\frac{\\pi^2}{6}
$$

$$
\\nabla \\times \\mathbf{E} = -\\frac{\\partial \\mathbf{B}}{\\partial t}
$$

### Horizontal Rule

---

### HTML in Markdown

<details>
<summary>Click to expand</summary>

This content is hidden by default. You can use HTML directly in your markdown files!

- Item 1
- Item 2
- Item 3

</details>

---

*Built with ❤️ using React, Vite, and TypeScript*
`;
