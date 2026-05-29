export async function normalizePptxZip(zip) {
  const contentTypes = zip.file('[Content_Types].xml');
  if (!contentTypes) return zip;

  const xml = await contentTypes.async('string');
  const normalized = xml.replace(/<Override[^>]+PartName="\/ppt\/fonts\/[^"]+"[^>]*\/>/g, '');
  zip.file('[Content_Types].xml', normalized);
  return zip;
}
