/* eslint-env node */

"use strict";

const octokit = require("./octokit-cache.js");

class Repository {
  constructor(name, ttl) {
    this.full_name = name;
    const parts = name.split('/');
    this.owner = parts[0];
    this.name = parts[1];
    this.ttl = ttl;
  }

  // retrieve and normalize w3c.json
  get w3c() {
    return octokit.get(`/extra/repos/${this.full_name}/w3c.json`).then(data => {
      if (data.group && !Array.isArray(data.group)) {
        data.group = [data.group];
      }
      return data;
    });
  }

  get config() {
    return octokit.get(`/v3/repos/${this.full_name}`)
      .then(data => {
        return this.w3c.then(w3c => {
          data.w3c = w3c;
          return data;
        });
      }).catch(() => {});
  }

  async createContent(path, message, content, branch) {
    let file = await octokit.request("GET /repos/:repo/contents/:path", {
      repo: this.full_name,
      path: path
    }).catch(err => {
      return err;
    });

    let sha;
    if (file.status === 200) {
      if (file.data.type !== "file") {
        throw new Error(`${path} isn't a file to be updated. it's ${file.data.type}.`);
      }
      // we're about to update the file
      sha = file.data.sha;
    } else if (file.status === 404) {
      // we're about to create the file
    } else {
      throw file;
    }
    content = Buffer.from(content, "utf-8").toString('base64');
    return octokit.request("PUT /repos/:repo/contents/:path", {
      repo: this.full_name,
      path: path,
      message: message,
      content: content,
      sha: sha,
      branch: branch
    });
  }

}

class GitHub {

  get ratelimit() {
    return octokit.request(`GET /rate_limit`).then(r => r.data);
  }

}

module.exports = { Repository, GitHub };
