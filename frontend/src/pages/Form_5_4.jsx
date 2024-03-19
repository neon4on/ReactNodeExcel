import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useCookies } from 'react-cookie';

const now = new Date();
const expires = new Date(now.getFullYear() + 10, now.getMonth(), now.getDate());

const Form54 = () => {
  const [cookies, setCookie] = useCookies(['tableData']);
  const [tableData, setTableData] = useState(
    () =>
      cookies.tableData || {
        winner: '',
        commandData1: '',
        commandData2: '',
        commandData3: '',
        select: '5.4.1',
      },
  );

  const handleChange = (e) => {
    const { name, value } = e.target;
    setTableData((prevState) => ({
      ...prevState,
      [name]: value,
    }));
    setCookie('tableData', { ...tableData, [name]: value }, { path: '/51', expires });
  };

  useEffect(() => {
    setTableData((prevTableData) => ({
      ...prevTableData,
      ...cookies.tableData,
    }));
  }, [cookies.tableData, setTableData]);

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      console.log('Отправляемые данные:', tableData);
      const response = await axios.post('http://localhost:4000/api/createExcel54', tableData);
      console.log('Ответ сервера:', response.data);
    } catch (error) {
      console.error('Ошибка при отправке данных:', error);
    }
  };

  return (
    <div className="container">
      <form onSubmit={handleSubmit}>
        <h5>
          Таблица "Количество победителей и призеров в командном или личном зачетах
          интеллектуальных, творческих конкурсов, иных мероприятий, направленных на развитие у
          обучающихся способностей в научной (научно-исследовательской), инженерно-технической,
          изобретательской и творческой сферах"
        </h5>

        <label className="labelSelect">
          Выберите тип соревнований:
          <select name="select" value={tableData.select || ''} onChange={handleChange} required>
            <option disabled value="DEFAULT">
              Не выбрано
            </option>
            <option value="5.4.1">Муниципальный</option>
            <option value="5.4.2">Региональный</option>
            <option value="5.4.3">Всероссийский</option>
            <option value="5.4.4">Международный</option>
          </select>
        </label>
        <textarea
          id="textarea1"
          className="materialize-textarea"
          type="text"
          name="winner"
          value={tableData.winner || ''}
          onChange={handleChange}
          rows="20"
          cols="20"
          wrap="hard"
          placeholder='Всероссийский конкурс "Большая перемена"'></textarea>
        <label htmlFor="textarea1">Название</label>
        <table className="iksweb">
          <tbody>
            <tr>
              <td>Побед</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData1"
                  value={tableData.commandData1 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td>Призовых мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData2"
                  value={tableData.commandData2 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td>Отсутствие соревновательной составляющей</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData3"
                  value={tableData.commandData3 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
          </tbody>
        </table>
        <div className="wrapper-button">
          <button className="btn waves-effect waves-light input-field" type="submit" name="action">
            Отправить
            <i className="material-icons right"></i>
          </button>
        </div>
      </form>
    </div>
  );
};

export default Form54;
