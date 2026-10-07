import App from './App';
import { DemoProvider } from './context';
export default function DemoEntry() { return <DemoProvider><App /></DemoProvider>; }
