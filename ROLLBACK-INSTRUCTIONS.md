# 🔄 ROLLBACK INSTRUCTIONS

**Created:** December 2, 2025  
**Before:** ChatGPT Architecture Migration (Non-Streaming)  
**Last Working Version:** v1.0.0.87 (Streaming)

---

## ⚠️ IF YOU NEED TO ROLLBACK

### Method 1: Git Tag (Recommended)

```bash
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised"

# Rollback to last working streaming version
git checkout v1.0.0.87-streaming-working

# Force push if needed (CAREFUL!)
git push origin main --force
```

### Method 2: Git Branch

```bash
cd "/Users/chriscorcoran/Documents/Cursor/NetSuite Formulas Revised"

# Switch to backup branch
git checkout backup-streaming-architecture

# Merge back to main if needed
git checkout main
git reset --hard backup-streaming-architecture
git push origin main --force
```

### Method 3: Local Backup Copy

```bash
# Find backup directory
ls -la "/Users/chriscorcoran/Documents/Cursor/" | grep "BACKUP"

# Copy back
cd "/Users/chriscorcoran/Documents/Cursor/"
rm -rf "NetSuite Formulas Revised"
cp -r "NetSuite Formulas Revised.BACKUP.YYYYMMDD-HHMMSS" "NetSuite Formulas Revised"
```

---

## 📦 BACKUP MANIFEST

### What Was Backed Up:

1. **Git Tag:** `v1.0.0.87-streaming-working`
   - Full git history
   - All files committed
   - Easy checkout

2. **Git Branch:** `backup-streaming-architecture`
   - Separate branch with working version
   - Can merge back if needed

3. **Local Copy:** `../NetSuite Formulas Revised.BACKUP.YYYYMMDD-HHMMSS/`
   - Complete filesystem copy
   - Not dependent on git
   - Can copy back manually

4. **GitHub:** All pushed to remote
   - Tag: `v1.0.0.87-streaming-working`
   - Branch: `backup-streaming-architecture`
   - Safe in cloud

---

## 🔍 VERIFY BACKUP

```bash
# Check tag exists
git tag -l | grep streaming-working

# Check branch exists
git branch -a | grep backup-streaming

# Check local backup
ls -la ../NetSuite\ Formulas\ Revised.BACKUP.*

# Check GitHub
git ls-remote --tags origin | grep streaming-working
```

---

## 📝 WHAT'S IN THE BACKUP

### Working Features (Streaming Version):

✅ All custom functions working:
   - NS.GLATITLE
   - NS.GLABAL
   - NS.GLABUD
   - NS.GLACCTTYPE
   - NS.GLAPARENT

✅ Streaming architecture:
   - Handles long queries (> 5 seconds)
   - Cancellation support
   - Robust error handling

✅ Task pane features:
   - Refresh All
   - Refresh Selected
   - Account search (by number and type)
   - Drill-down

✅ Backend:
   - Batch processing
   - BUILTIN.CONSOLIDATE
   - Separate P&L / Balance Sheet logic

✅ Known Issues:
   - Shows @ symbols on open (recalculation)
   - Performance could be better

### What Was Changing:

- Converting from streaming to non-streaming async
- Moving SuiteQL calls to task pane
- Implementing IndexedDB cache
- Adding smart prefetching

---

## 🚨 WHEN TO ROLLBACK

### Rollback If:

❌ Non-streaming functions timeout (> 5 seconds)  
❌ Formulas return #VALUE# or errors  
❌ Cache not working properly  
❌ IndexedDB issues in Excel  
❌ Task pane data engine broken  
❌ Any critical functionality lost  

### Don't Rollback If:

✅ @ symbols still show (optimization in progress)  
✅ Cache needs tuning (can fix)  
✅ Minor bugs (can debug)  
✅ Performance needs improvement (can optimize)  

---

## 📞 CHECKLIST BEFORE ROLLING BACK

1. **Document the issue:**
   - What's broken?
   - Error messages?
   - Steps to reproduce?

2. **Check console logs:**
   - Any JavaScript errors?
   - Network errors?
   - Cache errors?

3. **Try quick fixes first:**
   - Clear Excel cache
   - Restart Excel
   - Restart backend server
   - Clear browser cache (Cmd+Shift+R)

4. **If still broken:**
   - Proceed with rollback
   - Document what went wrong
   - Use for future debugging

---

## ✅ AFTER ROLLBACK

1. **Verify everything works:**
   ```
   =NS.GLATITLE(4220)                      → Works?
   =NS.GLABAL(4220,"Jan 2025","Jan 2025")  → Works?
   =NS.GLACCTTYPE(4220)                    → Works?
   ```

2. **Restart backend if needed:**
   ```bash
   ./restart-servers.sh
   ```

3. **Update Cloudflare Worker if tunnel changed**

4. **Test task pane features:**
   - Refresh All
   - Account search
   - Drill-down

---

## 📊 MIGRATION STATUS TRACKING

### Phase 1: Convert to Non-Streaming
- [ ] Started
- [ ] Functions.js updated
- [ ] Functions.json updated
- [ ] Tested
- [ ] Issues found (describe below)

### Phase 2: IndexedDB Cache
- [ ] Started
- [ ] Cache manager implemented
- [ ] Tested
- [ ] Issues found (describe below)

### Phase 3: Task Pane Data Engine
- [ ] Started
- [ ] Data engine implemented
- [ ] Tested
- [ ] Issues found (describe below)

### Phase 4: Smart Prefetching
- [ ] Started
- [ ] Prefetch logic implemented
- [ ] Tested
- [ ] Issues found (describe below)

---

## 📝 NOTES

Use this space to track issues during migration:

```
Date: ___________
Issue: ___________________________________________________________
Resolution: ______________________________________________________

Date: ___________
Issue: ___________________________________________________________
Resolution: ______________________________________________________
```

---

**Remember:** We have multiple backups. You can always rollback safely! 🛡️

