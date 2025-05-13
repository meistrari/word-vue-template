import { createApp } from 'vue';
import App from '@/app.vue';


Office.initialize = () => {
    console.log("Office initialized");
    const app = createApp(App);
    app.mount('#app');
};

// Render anyway outside of Office
const app = createApp(App);
app.mount('#app');