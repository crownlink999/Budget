// BMTS: приёмник файлов: ПК -> этот скрипт (script.google.com) -> GitHub input/
// Настройки хранятся в Project Settings -> Script properties:
//   GH_TOKEN  = токен GitHub (fine-grained, repo Balans_BMTS, Contents: Read and write)
//   KEY       = любой длинный пароль (тот же, что во 2-й строке send.txt)
//   REPO      = crownlink999/Balans_BMTS   (по умолчанию)
//   BRANCH    = main                       (по умолчанию)
function doGet() { return ContentService.createTextOutput('BMTS ok'); }

function doPost(e) {
  try {
    var p = PropertiesService.getScriptProperties();
    var body = JSON.parse(e.postData.contents);
    if (!p.getProperty('KEY') || body.key !== p.getProperty('KEY')) return out({ok:false, error:'bad key'});
    var repo = p.getProperty('REPO') || 'crownlink999/Balans_BMTS';
    var branch = p.getProperty('BRANCH') || 'main';
    var files = body.files || [];
    if (!files.length) return out({ok:false, error:'no files'});

    var base = 'https://api.github.com/repos/' + repo;
    var ref = gh(base + '/git/ref/heads/' + branch);
    var parent = ref.object.sha;
    var parentCommit = gh(base + '/git/commits/' + parent);
    var tree = [];
    files.forEach(function (f) {
      var name = String(f.name).replace(/[\\\/]/g, '_');
      var blob = gh(base + '/git/blobs', 'post', {content: f.b64, encoding: 'base64'});
      tree.push({path: 'input/' + name, mode: '100644', type: 'blob', sha: blob.sha});
    });
    var nt = gh(base + '/git/trees', 'post', {base_tree: parentCommit.tree.sha, tree: tree});
    var c = gh(base + '/git/commits', 'post', {message: 'Файлы с рабочего ПК: ' + files.length + ' шт.', tree: nt.sha, parents: [parent]});
    gh(base + '/git/refs/heads/' + branch, 'patch', {sha: c.sha});
    return out({ok:true, files: files.length, commit: c.sha});
  } catch (err) {
    return out({ok:false, error: String(err)});
  }
}

function gh(url, method, payload) {
  var opt = {method: method || 'get', muteHttpExceptions: true, contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + PropertiesService.getScriptProperties().getProperty('GH_TOKEN'),
              Accept: 'application/vnd.github+json'}};
  if (payload) opt.payload = JSON.stringify(payload);
  var r = UrlFetchApp.fetch(url, opt);
  if (r.getResponseCode() >= 300) throw new Error('GitHub ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
  return JSON.parse(r.getContentText());
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
