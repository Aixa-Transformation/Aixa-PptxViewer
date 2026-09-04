PPTX fidelity test for dfklgkjdfgj.pptx

Reference renderer: Microsoft PowerPoint Desktop
Comparison renderer: current Aixa PptxViewer core/library build

Primary confirmed issue:
- Slides 4-7, 10-11, and 13 contain a master shape with p:sp/@useBgFill="1".
- PowerPoint renders that shape using the white slide background.
- The current loader ignores useBgFill and resolves style fillRef idx="1" to Savon accent1 #94B6D2.

Files:
- slide-04-theme-fidelity-comparison.png: large side-by-side proof
- all-slides-fidelity-contact-sheet.png: complete 13-slide audit
- slide-04-difference-heatmap.png: pixel difference visualization
- fidelity-metrics.json: numeric differences for every slide