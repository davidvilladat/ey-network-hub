// Pages middleware — blocks public access to repository internals and source/build
// files. Added 2026-06-14 after the .git directory was found publicly downloadable.
// Static assets (HTML/CSS/JS/JSON/images) pass through unchanged.
export async function onRequest(context) {
  const { request, next } = context;
  const path = new URL(request.url).pathname.toLowerCase();
  const blocked =
    path.startsWith("/.git") ||
    path === "/.gitignore" ||
    path === "/.gitattributes" ||
    /\.(py|ps1|sh|bat|log|toml|lock|ipynb)$/.test(path);
  if (blocked) return new Response("Not found", { status: 404 });
  return next();
}
