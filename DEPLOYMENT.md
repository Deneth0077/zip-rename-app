# 🚀 Vercel Deployment Guide | Vercel වෙත Deploy කිරීමේ මඟපෙන්වීම

මෙම **ZIP Rename Nexus** යෙදුම Next.js මඟින් නිපදවා ඇති අතර එය 100% ක්ම පරිශීලකයාගේ බ්‍රවුසරය (Client-side) තුල ක්‍රියාත්මක වේ. එමනිසා මෙය කිසිදු වියදමකින් තොරව (Free Tier) ඉතා වේගවත්ව **Vercel** වෙත deploy කළ හැක.

Vercel වෙත පහසුවෙන්ම Deploy කරගැනීම සඳහා ක්‍රම දෙකක් (Methods) පහත දැක්වේ:

---

## 🛠️ Method 1: Connecting with GitHub (වඩාත් පහසු සහ නිර්දේශිත ක්‍රමය)

මෙය ඔබගේ Project එක GitHub ගිණුමට ඇතුලත් කර Vercel සමඟ සම්බන්ධ කරන වඩාත්ම ජනප්‍රිය සහ පහසු ක්‍රමයයි.

### Step 1: Create a GitHub Repository (GitHub ගිණුමට Upload කිරීම)
1. [GitHub (github.com)](https://github.com/) වෙත ගොස් ඔබගේ ගිණුමට ලොග් වන්න.
2. **New Repository** එකක් සාදන්න (උදා: `zippass-rename-app`).
3. ඔබගේ පරිගණකයේ Command Prompt / VS Code Terminal එකෙන් පහත විධානයන් (Commands) ක්‍රියාත්මක කර Project එක GitHub වෙත Upload කරන්න:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - ZIP Rename Nexus"
   git branch -M main
   git remote add origin <your-github-repo-url>
   git push -u origin main
   ```

### Step 2: Import to Vercel (Vercel වෙත සම්බන්ධ කිරීම)
1. [Vercel Dashboard (vercel.com)](https://vercel.com/) වෙත ගොස් ලොග් වන්න.
2. **"Add New"** ක්ලික් කර **"Project"** තෝරන්න.
3. ඔබගේ GitHub ගිණුම සම්බන්ධ කර එහි ඇති `zippass-rename-app` repository එක ඉදිරියෙන් ඇති **"Import"** ක්ලික් කරන්න.
4. **Configure Project** පිටුවේ වෙනස්කම් කිසිවක් කිරීමට අවශ්‍ය නැත (සියලුම Settings ස්වයංක්‍රීයව සකස් වේ).
5. **"Deploy"** බොත්තම ක්ලික් කරන්න.
6. විනාඩියක් ඇතුලත ඔබගේ වෙබ් අඩවිය සජීවීව (Live) ක්‍රියාත්මක වන අතර ඔබට නොමිලේම `.vercel.app` domain එකක් ලැබෙනු ඇත!

---

## 💻 Method 2: Using Vercel CLI (VS Code Terminal එක භාවිතයෙන්)

ඔබට GitHub භාවිතා නොකර සෘජුවම VS Code Terminal එක මඟින් Deploy කිරීමට අවශ්‍ය නම් පහත පියවර අනුගමනය කරන්න.

### Step 1: Install Vercel CLI
Terminal එක තුල පහත විධානය ක්‍රියාත්මක කර Vercel CLI ස්ථාපනය කරන්න:
```bash
npm install -g vercel
```

### Step 2: Login & Deploy
1. Terminal එකේ පහත විධානය ක්‍රියාත්මක කරන්න:
   ```bash
   vercel
   ```
2. ඔබෙන් අසන ප්‍රශ්න වලට පහත පරිදි පිළිතුරු දෙන්න:
   - **Set up and deploy?** `yes` (Y)
   - **Which scope?** (ඔබගේ Vercel account එක තෝරන්න)
   - **Link to existing project?** `no` (N)
   - **What’s your project’s name?** `zippass-rename-app` (හෝ කැමති නමක්)
   - **In which directory is your code located?** `./` (එලෙසම Enter කරන්න)
   - **Want to modify build settings?** `no` (N)

3. මෙම ක්‍රියාවලිය අවසන් වූ පසු, ඔබට **Preview URL** එකක් ලැබෙනු ඇත.
4. අවසාන වශයෙන්, නිෂ්පාදන මට්ටමින් (Production) සජීවී කිරීමට පහත විධානය ලබා දෙන්න:
   ```bash
   vercel --prod
   ```

---

## 💡 වැදගත් කරුණු / Important Tips

* **Environment Variables:** මෙම ඇප් එක 100% ක්ම බ්‍රවුසරය තුල ක්‍රියාත්මක වන නිසා කිසිදු API Keys හෝ database credentials (Environment Variables) Vercel හි සැකසීමට අවශ්‍ය නොවේ!
* **Automatic Redeploy:** GitHub ක්‍රමය (Method 1) භාවිතා කළහොත්, ඔබ Project එකෙහි කරන ඕනෑම වෙනස්කමක් GitHub වෙත Push කළ සැනින් Vercel වෙබ් අඩවිය ස්වයංක්‍රීයව අලුත් (Auto-redeploy) වේ.
