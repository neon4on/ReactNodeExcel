import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useCookies } from 'react-cookie';

const now = new Date();
const expires = new Date(now.getFullYear() + 10, now.getMonth(), now.getDate());

const Form51 = () => {
  const [cookies, setCookie] = useCookies(['tableData']);
  const [tableData, setTableData] = useState(
    () =>
      cookies.tableData || {
        winner: '',
        commandData1: '',
        commandData2: '',
        commandData3: '',
        commandData11: '',
        commandData21: '',
        commandData31: '',
        personalData1: '',
        personalData2: '',
        personalData3: '',
        grandPrizeData: '',
        individualAchievementData: '',
        specialAwardsData: '',
        lackOfCompetitiveComponentData: '',
      },
  );

  const handleChange = (e) => {
    const { name, value } = e.target;
    setTableData((prevState) => ({
      ...prevState,
      [name]: value,
    }));
    setCookie('tableData', { ...tableData, [name]: value }, { path: '/52', expires });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      console.log('Отправляемые данные:', tableData);
      const response = await axios.post('http://localhost:4000/api/createExcel52', tableData);
      console.log('Ответ сервера:', response.data);
    } catch (error) {
      console.error('Ошибка при отправке данных:', error);
    }
  };

  useEffect(() => {
    setTableData((prevTableData) => ({
      ...prevTableData,
      ...cookies.tableData,
    }));
  }, [cookies.tableData, setTableData]);

  return (
    <div className="container">
      <form onSubmit={handleSubmit}>
        <h5>
          Таблица "Участие в международных и всеармейских тематических конкурсах, фестивалях,
          смотрах и других мероприятиях, проводимых по плану Минобороны России органами военного
          управления"
        </h5>
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
          placeholder="Всеармейский чемпионат по шахматам среди обучающихся довузовских образовательных организаций Минобороны Россия"></textarea>
        <label htmlFor="textarea1">Название</label>
        <table className="iksweb">
          <tbody>
            <tr>
              <td rowSpan="3">Призовые места по итогам командного первенства в номинациях</td>
              <td className="center-align">1-х мест</td>
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
              <td className="center-align">2-х мест</td>
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
              <td className="center-align">3-х мест</td>
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
            <tr>
              <td rowSpan="3">Призовые места по итогам командного первенства в номинациях</td>
              <td className="center-align">1-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData11"
                  value={tableData.commandData11 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td className="center-align">2-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData21"
                  value={tableData.commandData21 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td className="center-align">3-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="commandData31"
                  value={tableData.commandData31 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td rowSpan="3">Призовые места по итогам личного первенства в номинациях</td>
              <td className="center-align">1-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="personalData1"
                  value={tableData.personalData1 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td className="center-align">2-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="personalData2"
                  value={tableData.personalData2 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td className="center-align">3-х мест</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="personalData3"
                  value={tableData.personalData3 || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td colSpan="2">Гран При</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="grandPrizeData"
                  value={tableData.grandPrizeData || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td colSpan="2">Приз за отдельные достижения</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="individualAchievementData"
                  value={tableData.individualAchievementData || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td colSpan="2">Побед в специальных номинациях</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="specialAwardsData"
                  value={tableData.specialAwardsData || ''}
                  onChange={handleChange}
                />
              </td>
            </tr>
            <tr>
              <td colSpan="2">Отсутствие соревновательной составляющей</td>
              <td>
                <input
                  type="text"
                  className="input-field col s6"
                  name="lackOfCompetitiveComponentData"
                  value={tableData.lackOfCompetitiveComponentData || ''}
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

export default Form51;
