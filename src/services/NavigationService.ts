/**
 * NavigationService - Управління навігацією між сторінками
 * FIXED: Винесено з main.ts (hashchange обробники)
 * 
 * Відповідальність:
 * - Роутинг між сторінками додатку
 * - Обробка hash змін URL
 * - Показ/приховування секцій за маршрутом
 * - Виклик callback-ів при зміні маршруту
 * 
 * @class NavigationService
 */

import type { Route } from '../types';

type RouteCallback = (route: Route) => void;

export class NavigationService {
  private callbacks: RouteCallback[] = [];

  /**
   * Конструктор - ініціалізує NavigationService
   * FIXED: Налаштовує слухача hashchange
   */
  constructor() {
    this.setupHashChangeListener();
    this.handleInitialRoute();
  }

  /**
   * Налаштування слухача зміни hash в URL
   * FIXED: Підписка на hashchange події
   * @private
   */
  private setupHashChangeListener(): void {
    window.addEventListener('hashchange', () => {
      this.handleRouteChange();
    });
  }

  /**
   * Обробка початкового маршруту при завантаженні
   * @private
   */
  private handleInitialRoute(): void {
    const route = this.getCurrentRoute();
    this.showRoute(route);
    this.notifyCallbacks(route);
  }

  /**
   * Обробка зміни маршруту
   * @private
   */
  private handleRouteChange(): void {
    const route = this.getCurrentRoute();
    this.showRoute(route);
    this.notifyCallbacks(route);
  }

  /**
   * Отримання поточного маршруту з URL
   * @public
   */
  public getCurrentRoute(): Route {
    const hash = location.hash.slice(1) || '/functions';
    return hash as Route;
  }

  /**
   * Показ секції за маршрутом
   * @private
   */
  private showRoute(route: Route): void {
    // Ховаємо всі маршрути
    document.querySelectorAll<HTMLElement>('.route').forEach(el => {
      el.hidden = true;
    });

    // Показуємо потрібний маршрут
    const routeElement = document.querySelector<HTMLElement>(`[data-route="${route}"]`);
    if (routeElement) {
      routeElement.hidden = false;
    }
    
    // Оновлюємо активні пункти навігації
    document.querySelectorAll('.nav a').forEach(link => {
      const href = link.getAttribute('href') || '';
      link.classList.toggle('active', href === `#${route}`);
    });
  }

  /**
   * Навігація на інший маршрут
   * FIXED: Змінює hash та оновлює UI
   * @public
   */
  public navigateTo(route: Route): void {
    location.hash = `#/${route}`;
  }

  /**
   * Реєстрація callback для викликів при зміні маршруту
   * FIXED: Дозволяє іншим модулям реагувати на навігацію
   * @public
   */
  public onRouteChange(callback: RouteCallback): void {
    this.callbacks.push(callback);
  }

  /**
   * Виклик всіх зареєстрованих callbacks
   * FIXED: Сповіщає підписників про зміну маршруту
   * @private
   */
  private notifyCallbacks(route: Route): void {
    this.callbacks.forEach(callback => {
      try {
        callback(route);
      } catch (error) {
        console.error('Error in route callback:', error);
      }
    });
  }
}
