# !/bin/bash

local_github_maven_repo='/Users/MZ02-SEUNGPILPARK/IdeaProjects/MegazoneDSG/maven-repo'

./gradlew publishToMavenLocal
cp -R ~/.m2/repository/com/mz/poi-mapper ${local_github_maven_repo}/snapshots/com/mz
cd ${local_github_maven_repo}
git add ./
git commit -m "poi-mapper deploy"
git push origin master


