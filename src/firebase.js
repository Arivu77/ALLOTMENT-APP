import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyBYL-q4Y09AFDo1yaBBGGGfc15saJArx28",
  authDomain: "allotment-656e2.firebaseapp.com",
  databaseURL: "https://allotment-656e2-default-rtdb.firebaseio.com",
  projectId: "allotment-656e2",
  storageBucket: "allotment-656e2.firebasestorage.app",
  messagingSenderId: "1047387619531",
  appId: "1:1047387619531:web:e358295b31a2ca2bde1b4e",
  measurementId: "G-RV8WWPB23K"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();
export const db = getFirestore(app);
