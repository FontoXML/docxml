import { render } from 'preact';
import { App } from './App.tsx';
import './style.css';

const root = document.querySelector('#app');

if (!(root instanceof HTMLElement)) {
	throw new Error('Could not find #app root.');
}

render(<App />, root);
