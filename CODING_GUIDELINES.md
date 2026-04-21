# Developer & AI Coding Guidelines

This document serves as the primary rulebook for all developers and AI coding tools (Antigravity, Cursor, etc.) working on the **HSEQ-SWF** project. Following these rules ensures modularity, scalability, and seamless collaboration across multiple developers.

---

## 🏗️ 1. Project Architecture (Modular Design)

All new features must follow the **Modular Architecture**. 

### Root Structure:
- `src/modules/`: **DO NOT** put business logic directly inside the webpart folder. Each developer works in their own module folder.
  - `src/modules/projects/`
  - `src/modules/compliance/`
  - `src/modules/incidents/`
- `src/common/`: Shared code used by multiple modules.
  - `src/common/components/`: Reusable UI elements (Buttons, Tables, etc.).
  - `src/common/services/`: Centralized PnPjs services and common API logic.
  - `src/common/models/`: Global interfaces and types.
- `src/webparts/swf/`: The "App Shell". It should only handle routing, layout (Sidebar/Header), and module switching.

---

## ⚛️ 2. React & TypeScript Standards

- **Functional Components**: Always use functional components with `React.FC` or standard function declarations.
- **Strong Typing**: Avoid `any`. Define interfaces for all Props, State, and API Responses in a `models/` directory.
- **Separation of Concerns**: 
  - Keep components visual-only where possible.
  - Move complex logic/state into **Custom Hooks** (`hooks/`).
  - Move data fetching/API calls into **Services** (`services/`).
- **File Size**: If a component file exceeds **300 lines**, it must be refactored into smaller sub-components or hooks.

---

## 📊 3. Data Fetching (PnPjs)

- All SharePoint interactions must use **@pnp/sp**.
- Use the centralized PnPjs configuration.
- Shared services should be located in `src/common/services/SharePointService.ts`.
- Module-specific services should be in `src/modules/[module]/services/`.

---

## 🎨 4. Styling Guidelines

- **SCSS Modules**: Use `[ComponentName].module.scss` for component-specific styling.
- **Theming**: Use the project's global CSS variables and SPFx theme tokens to support Dark/Light mode consistently.
- **Consistency**: Before creating a new style, check if a shared UI component in `src/common/components` already exists.

---

## 🤖 5. Instructions for AI Assistants (Antigravity/Cursor)

> [!TIP]
> AI agents should read these rules at the start of every task.

1. **Context Awareness**: Before generating code, analyze `src/common` to see if the required utility or component already exists.
2. **Path Compliance**: Always place new feature code in `src/modules/[module_name]`.
3. **Refactoring**: If asked to "add a feature" to a large existing component, proactively suggest moving logic to a custom hook to maintain readability.
4. **Consistency**: Ensure all generated code matches the existing naming conventions (PascalCase for components, camelCase for variables/functions).

---

## 🛡️ 6. Git Protocol

- Work on feature branches: `feature/[developer_name]/[module_name]`.
- Keep PRs small and module-specific.
- Always run `gulp build` to ensure no linting or TypeScript errors before pushing.
