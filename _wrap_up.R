Sys.sleep(1)

# Rename the markdown file
if (file.exists("README.html.md")) {
  file.rename("README.html.md", "README.md")
}
if (file.exists("README.html")) {
  file.rename("README.html", "index.html")
}
