# Electrobaza

A mobile-first library of lectures, practical exercises, laboratory work, and presentations. The site is a static GitHub Pages project: students do not need accounts, and no separate server is required.

## Publish with GitHub Pages

1. Upload the project files and the `content` folder to the repository root.
2. Open **Settings → Pages** in the GitHub repository.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select branch **main**, folder **/(root)**, and click **Save**.
5. Wait for the deployment to finish in **Actions**. The published address will appear in the Pages settings.

GitHub Pages republishes the site after each commit to `main`.

## Project files

- `index.html` — site entry point.
- `styles.css` — layout and mobile typography.
- `app.js` — navigation between groups, subjects, topics, and materials.
- `materials.json` — the catalogue of groups, topics, and material entries.
- `content/` — HTML fragments for lessons, assignments, and slides.
- `catalogue-data.js` — generated local copy of the catalogue and lesson HTML, so `index.html` also works when opened directly as a file.
- `build-catalogue.mjs` / `build-catalogue.cmd` — regenerate that local copy.

## Open locally

Double-click `index.html` for a simple local preview. After changing `materials.json` or a file in `content/`, run `build-catalogue.cmd`, then refresh the page. For a live preview that rebuilds the catalogue and refreshes the browser after each saved change, run `preview.cmd` instead. Keep its console window open while previewing.

## Add a material

For the detailed Russian step-by-step guide with lecture and practical-work templates, see [ИНСТРУКЦИЯ_ПО_ДОБАВЛЕНИЮ_МАТЕРИАЛОВ.md](ИНСТРУКЦИЯ_ПО_ДОБАВЛЕНИЮ_МАТЕРИАЛОВ.md).

1. Create an HTML file in `content/`, such as `content/new-lesson.html`. Add only lesson content: headings, paragraphs, lists, images, and formulas.
2. Add an entry to the `materials` array in `materials.json`. For example:

   ```json
   {"id":"new-lesson","type":"lecture","name":"Лекция","description":"Краткое описание","file":"content/groups/1eo-25/mdk-04-02/semester-1/<topic-id>/new-lesson.html"}
   ```

3. Add the material ID to the `materials` list of the relevant topic.
4. To share one material with several groups, add its topic to each group's subject catalogue. Reuse the same material ID and HTML file.
5. Run `build-catalogue.cmd` before publishing so `catalogue-data.js` contains the updated catalogue and HTML.
6. In GitHub Desktop, review the changed files, commit them to `main`, and click **Push origin**.

Supported types are `lecture`, `practice`, `lab`, `slides`, and `file`. The catalogue must remain valid JSON. The text, materials, and repository are public; do not add students' personal data or restricted content.

## Reading and media

HTML fragments are displayed inside the same responsive reader and use the site's common typography. The default text is large and can be increased by the reader. Store course material under `content/groups/<group-id>/<discipline-id>/<topic-id>/`; add a `semester-<number>/` folder inside a discipline when that group's course variant depends on the semester. Keep images beside the material and reference them with a relative path.
