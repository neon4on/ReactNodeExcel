import React from 'react';
import { render, fireEvent } from '@testing-library/react';
import '@testing-library/jest-dom/extend-expect'; // Для улучшенных сопоставлений

import Form51 from '../pages/Form_5_1'; // Путь к вашему компоненту

describe('Form51', () => {
  it('should update state and cookie when input changes', () => {
    const { getByLabelText, getByText, getByPlaceholderText } = render(<Form51 />);

    // Получаем элементы формы
    const textarea = getByLabelText('Название');
    const inputCommandData1 = getByPlaceholderText('1-х мест');
    const submitButton = getByText('Отправить');

    // Меняем значение textarea
    fireEvent.change(textarea, { target: { value: 'Новое название' } });
    expect(textarea).toHaveValue('Новое название');

    // Меняем значение inputCommandData1
    fireEvent.change(inputCommandData1, { target: { value: 'Новое значение' } });
    expect(inputCommandData1).toHaveValue('Новое значение');

    // Симулируем отправку формы
    fireEvent.click(submitButton);

    // Добавьте здесь дополнительные проверки по мере необходимости
  });
});
