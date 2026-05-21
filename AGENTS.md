# Repository Guidelines

## Project Structure & Module Organization

This repository is a Vite React application for a motorcycle parts UI. The main application code lives in `src/`: `main.jsx` mounts the app, `App.jsx` contains the primary interface, and `App.css` / `index.css` hold styling. Image and SVG assets imported by React belong in `src/assets/`. Static files served directly by Vite, such as `favicon.svg` and sprite files, live in `public/`. Build output is generated in `dist/` and should not be edited or committed.

## Build, Test, and Development Commands

- `npm install`: install dependencies from `package-lock.json`.
- `npm run dev`: start the Vite development server with hot reload.
- `npm run build`: create a production build in `dist/`.
- `npm run preview`: serve the production build locally for verification.
- `npm run lint`: run ESLint across the repository.

Run commands from the repository root.

## Coding Style & Naming Conventions

Use modern ES modules and React function components. Keep component names in PascalCase, for example `ProductCard`, and use camelCase for variables, functions, props, and event handlers. Use `.jsx` for files containing JSX and `.js` for plain JavaScript utilities. Match the existing style: two-space indentation, single quotes, and no semicolons. Keep component-specific styles close to the component stylesheet already used by the app unless a broader shared style is needed.

## Testing Guidelines

No automated test framework is currently configured. Before opening a change, run `npm run lint` and `npm run build`. For UI changes, also run `npm run dev` and manually verify the affected flow in a browser. If tests are added later, prefer colocated names such as `ComponentName.test.jsx` or a `src/__tests__/` directory, and document the new test command in `package.json`.

## Commit & Pull Request Guidelines

This repository has no commit history yet, so use a simple, consistent convention going forward: short imperative subjects such as `Add product search filter` or `Fix mobile header spacing`. Keep commits focused on one logical change.

Pull requests should include a concise summary, manual verification steps, and screenshots or short recordings for visible UI changes. Link related issues when available. Confirm `npm run lint` and `npm run build` before requesting review.

## Agent-Specific Instructions

Preserve existing project patterns and keep changes scoped. Do not edit generated output in `dist/` or dependency files in `node_modules/`. Prefer small React components and clear data structures over broad rewrites.
