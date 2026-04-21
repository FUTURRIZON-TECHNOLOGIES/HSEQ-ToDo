# 🚀 Module Migration Guide (Old Style to New Modular)

If you have been working on a module in the old monolithic structure (where everything was inside the webpart folder), follow this guide to move your work into the new **Zero-Conflict** architecture.

---

## 🛠️ Step 1: The New Setup
1. **Pull the latest code**: Ensure you have the `src/modules` and `src/common` directories.
2. **Review `CODING_GUIDELINES.md`**: Understand the new folder roles.

---

## 📦 Step 2: Move your UI Components
**Old Location:** `src/webparts/swf/components/Modules/[YourModule]`  
**New Location:** `src/modules/[your-module-name]`

1. Create your module folder: `src/modules/[your-module-name]`.
2. Move your components, styles, and hooks there.
3. **Update Imports:** Your relative paths will change. Use the AI tool to fix them (see prompt below).

---

## ⚙️ Step 3: Extract your Service Logic
**Old Location:** `src/webparts/swf/services/SPService.ts`  
**New Location:** `src/modules/[your-module-name]/services/[YourService].ts`

1. Create a service file in your module folder.
2. Copy **ONLY** your feature-specific methods (e.g., `getProjects`, `addProject`) from the old `SPService.ts`.
3. Have your new service extend `BaseSPService` from `src/common/services/BaseSPService`.

---

## 📝 Step 4: Register your Module
Open `src/common/config/ModuleRegistry.tsx` and add your module to the list:

```typescript
export const ModuleRegistry: IModuleInfo[] = [
    // ... existing modules
    {
        id: 'MyModule',
        label: 'My New Module',
        iconName: 'DocumentSearch', // Fluent UI Icon Name
        group: 'HSEQ',
        component: MyModuleMainComponent
    }
];
```

---

## 🤖 Step 5: Using AI to do the work
Copy and paste this prompt into **Antigravity**, **Claude Code**, or **Cursor**:

> "I need to migrate my module '[Name]' to the new architecture. 
> 1. Create a folder `src/modules/[name]`.
> 2. Move my components from the old webpart folder to the new modules folder.
> 3. Extract my methods from `SPService.ts` and create a dedicated service in my module folder that extends `BaseSPService`.
> 4. Fix all broken imports.
> 5. Add my module to `src/common/config/ModuleRegistry.tsx`.
> Follow the `CODING_GUIDELINES.md` strictly."

---

## ✅ Benefits
- You will **never** have merge conflicts in the Sidebar or WebPart again.
- Your code is completely isolated.
- The App Shell handles all the routing and layout for you automatically.
