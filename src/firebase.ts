import { initializeApp } from "firebase/app";
import {
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager,
  setLogLevel,
} from "firebase/firestore";
import firebaseConfigData from "../firebase-applet-config.json";

// Log only errors to keep console clean while surfacing real sync issues
setLogLevel('error');

const firebaseConfig = {
  apiKey: firebaseConfigData.apiKey,
  authDomain: firebaseConfigData.authDomain,
  projectId: firebaseConfigData.projectId,
  storageBucket: firebaseConfigData.storageBucket,
  messagingSenderId: firebaseConfigData.messagingSenderId,
  appId: firebaseConfigData.appId,
};

const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(
  app,
  {
    experimentalAutoDetectLongPolling: true,
    localCache: persistentLocalCache({
      tabManager: persistentMultipleTabManager(),
    }),
  },
  firebaseConfigData.firestoreDatabaseId
);
