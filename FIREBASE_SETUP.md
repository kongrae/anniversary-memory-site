# Firebase Storage setup

1. Firebase console에서 Web app을 만들고 `firebaseConfig` 값을 복사합니다.
2. `index.html`의 `firebaseConfig` 빈 문자열을 복사한 값으로 채웁니다.
3. Authentication에서 Anonymous sign-in을 켭니다.
4. Storage Rules에 `firebase.storage.rules` 내용을 붙여넣고 게시합니다.
5. 배포 후 가운데 기념일 사진을 선택하면 `anniversary/current-main.jpg`로 업로드됩니다.

업로드된 사진은 Firebase Storage URL로 다시 로드되므로 다른 휴대폰에서도 같은 사진이 보입니다.
