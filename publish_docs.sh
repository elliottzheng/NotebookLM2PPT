bun run docs:build
tmp_folder=../temp-gh-pages
# abs path of tmp_folder
tmp_folder=$(pwd)/$tmp_folder
echo $tmp_folder
rm -rf $tmp_folder
mkdir $tmp_folder
cp -r .vitepress/dist/* $tmp_folder
cd $tmp_folder
git init
touch .nojekyll # disable jekyll so that github pages can render the site correctly
git add .
git config user.name "GitHub Actions"
git config user.email "actions@github.com"
git commit -m "docs: deploy"
git remote add origin https://github.com/elliottzheng/NotebookLM2PPT.git 2>nul || git remote set-url origin https://github.com/elliottzheng/NotebookLM2PPT.git
git push -f origin main:gh-pages
cd ..
