If you don't want to use special characters like pipes (`|`), there are a few other ways we can architect this so the agent still builds dynamic tables automatically.

Here are the best alternatives, ordered from easiest to most advanced:

### Option 1: Auto-Detect "Key-Value" Pairs (Easiest for Data Entry)

If your data naturally looks like a list with colons (`:`) or hyphens (`-`), you can just type it normally in Excel:

<pre><div node="[object Object]" class="relative whitespace-pre-wrap word-break-all my-2 rounded-xl bg-muted border"><div class="min-h-7 relative box-border flex flex-row items-center justify-between rounded-t border-b border-border px-2 py-0.5"><div class="font-sans text-sm text-muted-foreground">text</div><div class="flex flex-row gap-2 justify-end"></div></div><div class="p-3"><div class="w-full h-full text-xs cursor-text"><div class="code-block"><div class="code-line" data-line-number="1" data-line-start="1" data-line-end="1"><div class="line-content"><span class="mtk1">Area: Amendment</span></div></div><div class="code-line" data-line-number="2" data-line-start="2" data-line-end="2"><div class="line-content"><span class="mtk1">HVDLE Threshold: Monetary threshold for classification...</span></div></div><div class="code-line" data-line-number="3" data-line-start="3" data-line-end="3"><div class="line-content"><span class="mtk1">Applicability of CG Norms: Corporate Governance provisions...</span></div></div></div></div></div></div></pre>

**How it works:** I can program the agent to look for consecutive lines that contain a colon `:`. When it sees a block of these, it automatically splits the text at the colon. The left side becomes Column 1 ("Area"), and the right side becomes Column 2 ("Amendment"). It then draws the table automatically.

### Option 2: Dedicated Table Columns in Excel

Instead of forcing everything into the `Summary` column, we can add two new columns to your Excel template: `Table_Keys` and `Table_Values`. **How it works:**

* In `Table_Keys`, you press Alt+Enter to list the areas: `HVDLE Threshold` `Applicability of CG Norms`
* In `Table_Values`, you list the amendments.
* If these columns have data, the agent reads them side-by-side and generates a table on the slide.

### Option 3: AI-Powered Table Extraction (Most Advanced)

Since this is an AI agent, we can integrate an LLM (like Gemini or OpenAI) directly into `agent.py`. **How it works:** You type your summary in plain English, however you want. The script sends the raw text to the LLM and asks, *"Is there tabular data hidden in this text? If so, extract it into a table."* The LLM parses it, and the Python script uses that structured data to draw the PowerPoint table.

**Which of these sounds best for your workflow?** Option 1 is usually the crowd-favorite because you just type naturally with colons, and the code does all the heavy lifting!

4:36 PM, 6/12/2026
