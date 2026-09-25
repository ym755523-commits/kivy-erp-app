[app]
title = Test Customs
package.name = testcustoms
package.domain = org.test

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,ttf,db,json

version = 0.1

requirements = python3,kivy

orientation = portrait
fullscreen = 0

android.api = 33
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a

android.permissions = INTERNET

[buildozer]
log_level = 2
warn_on_root = 1
