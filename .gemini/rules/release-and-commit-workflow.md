# Release, Versioning and Commit Workflow Rules

This document specifies the standard, mandatory procedure for preparing commits, versioning, updating release notes, and publishing packages to GitHub and NuGet for the **Conq.Noble.Spreadox** project.

---

## 1. Golden Rules of Changes & Versioning

1. **Clean Separation of Scopes:**
   - Never combine unrelated features, bug fixes, or test suites into a single release or single commit.
   - Always run `git status` to distinguish between staged, unstaged, and untracked changes.
   - If work spans multiple release milestones (e.g. baseline QA hardening vs. advanced fuzzing/resilience), commit and release them as sequential, self-contained versions (e.g., `v0.9.1` then `v0.9.2`).

2. **Semantic Versioning 2.0 (SemVer):**
   - **Patch (`0.9.X`):** Backwards-compatible bug fixes, performance optimizations, memory leak fixes, internal refactoring, and test additions.
   - **Minor (`0.X.0`):** New backwards-compatible public features, new chart types, or export formats.
   - **Major (`X.0.0`):** Breaking changes to the public API surface. *(Note: Spreadox maintains a strict policy to preserve the public API without breaking changes).*

3. **Single Source of Truth for Version:**
   - The version must be declared exclusively in `Directory.Build.props`:
     ```xml
     <Version>X.Y.Z</Version>
     <AssemblyVersion>X.Y.Z.0</AssemblyVersion>
     <FileVersion>X.Y.Z.0</FileVersion>
     <InformationalVersion>X.Y.Z</InformationalVersion>
     ```
   - Never hardcode versions in individual `.csproj` files.

4. **Changelog Integrity (`CHANGELOG.md`):**
   - Every release commit MUST update `CHANGELOG.md` following [Keep a Changelog](https://keepachangelog.com/).
   - Standard subsections: `### Added`, `### Changed`, `### Fixed`, `### Performance`, `### Security`.
   - Include version header with ISO 8601 release date: `## [X.Y.Z] - YYYY-MM-DD`.

---

## 2. Conventional Commit Standards

Every commit message must follow the Conventional Commits specification:

```
<type>(<scope>): <concise imperative summary>

[Optional detailed body explaining WHY the change was made, invariant details, and metrics]

- Key change 1
- Key change 2
```

### Commit Types:
- `feat`: A new user-facing feature or API capability.
- `fix`: A bug fix or boundary validation.
- `perf`: A code change that improves performance or reduces memory allocations without changing behavior.
- `test`: Adding missing tests, benchmarks, or correcting existing tests.
- `docs`: Documentation updates, benchmarks, and test results.
- `chore` / `ci`: Updating build configurations, CI workflows, or release automation.

### Standard Scopes:
`excel`, `pdf`, `streaming`, `charts`, `styling`, `core`, `qa`, `packaging`.

---

## 3. Pre-Commit Verification Gate

Before staging and committing any release:

```bash
# 1. Clean release build
dotnet build -c Release

# 2. Run automated tests (excluding long-running benchmarks if fast-cycle)
dotnet test Conq.Noble.Spreadox.sln -c Release --filter Category!=Perf

# 3. Local package verification
dotnet pack src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj -c Release -o artifacts

# 4. Verify package contents
# Ensure .nupkg and .snupkg are generated with XML docs and symbols included.
```

---

## 4. Step-by-Step Release Execution

When releasing a staged or unstaged milestone:

### Step 1: Stage Atomic Changes
```bash
# Stage the exact source files, tests, props, and changelog for the target version
git add <target-files> Directory.Build.props CHANGELOG.md
```

### Step 2: Commit with Conventional Commit Message
```bash
git commit -m "feat(scope): concise title" -m "Detailed description..."
```

### Step 3: Create Annotated Git Tag
The GitHub Actions release workflow (`.github/workflows/release.yml`) triggers exclusively on `v*.*.*` tags:
```bash
git tag -a vX.Y.Z -m "Release vX.Y.Z: Summary of changes"
```

### Step 4: Push to GitHub & Trigger NuGet Publish
```bash
# Push commit to main branch
git push origin main

# Push release tag to trigger automated build, test, pack, and keyless Trusted Publishing (OIDC via NuGet/login@v1) to nuget.org
git push origin vX.Y.Z
```

### Step 5: Verify Deployment
- Monitor GitHub Actions: `https://github.com/FatihBaycu/Conq.Noble.Spreadox/actions`
- Verify NuGet.org package: `https://www.nuget.org/packages/Conq.Noble.Spreadox/X.Y.Z`
