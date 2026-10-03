import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';

const pet = document.getElementById('pet') as HTMLDivElement;

function blink(): void {
  pet.classList.add('blink');
  setTimeout(() => pet.classList.remove('blink'), 150);
}

function scheduleBlink(): void {
  const delay = 1600 + Math.random() * 3400;
  setTimeout(() => {
    blink();
    scheduleBlink();
  }, delay);
}

scheduleBlink();

// 单击桌宠:唤起并聚焦主窗口(按住拖动则由 data-tauri-drag-region 处理,不会触发)
pet.addEventListener('click', () => {
  invoke('show_main');
  pet.classList.add('poke');
  setTimeout(() => pet.classList.remove('poke'), 320);
});

// 转换进行中:桌宠进入"忙碌"状态
listen<{ busy: boolean }>('convert-state', (e) => {
  pet.classList.toggle('working', !!e.payload?.busy);
});
