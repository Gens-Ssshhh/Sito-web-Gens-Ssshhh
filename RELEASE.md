# Gens Ssshhh release

- Release date: 2026-09-11
- Release id: 20260911
- Android version name: 1.0.20260911
- Android version code: 20260911
- Download artifact: `Gens-Ssshhh-App-YYYY-MM-DD_HH-MM/Gens-Ssshhh-App-YYYY-MM-DD_HH-MM.apk`

The Android workflow adds the Rome/Italy release time (`HH-MM`) to the APK filename and artifact name.
The Android workflow stamps the APK's bundled `www/release.json` with that build's version.
The committed `www/release.json` describes the newest APK actually uploaded to the shared Drive folder. Update its `version`, `releaseDate`, and `releaseId` only after uploading a newer APK to Drive, then deploy GitHub Pages. Web deployments do not change the app release metadata.
