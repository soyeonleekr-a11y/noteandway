import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

// Firebase 콘솔 > 프로젝트 설정 > 일반 > "내 앱" 에서 복사한 값을 그대로 붙여넣으세요.
// (이 값들은 클라이언트에 노출돼도 되는 공개 식별자입니다 — 실제 보안은 Firestore 규칙이 담당해요.)
const firebaseConfig = {
  apiKey: "AIzaSyCR5f0iPdaNB9e6sKo8Mj0Ww-JLU3WMOMw",
  authDomain: "noteandway-9ae5b.firebaseapp.com",
  projectId: "noteandway-9ae5b",
  storageBucket: "noteandway-9ae5b.firebasestorage.app",
  messagingSenderId: "984180791027",
  appId: "1:984180791027:web:d7a685179bcb6fc7974986",
};

export const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
