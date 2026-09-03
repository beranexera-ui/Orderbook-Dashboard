import { initializeApp } from "firebase/app";
import { initializeFirestore, setLogLevel } from "firebase/firestore";
import firebaseConfigData from "../firebase-applet-config.json";

// Suppress Firestore internal connection warnings in preview environment
setLogLevel('silent');

const firebaseConfig = {
  apiKey: firebaseConfigData.apiKey,
  authDomain: firebaseConfigData.authDomain,
  projectId: firebaseConfigData.projectId,
  storageBucket: firebaseConfigData.storageBucket,
  messagingSenderId: firebaseConfigData.messagingSenderId,
  appId: firebaseConfigData.appId,
};

const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(app, {
  experimentalForceLongPolling: true
}, firebaseConfigData.firestoreDatabaseId);
