import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
from pathlib import Path
from auth import check_password
import re
import yaml

# Настройка страницы
st.set_page_config(page_title="Анализ образовательных программ", layout="wide")

# Проверка авторизации
if not check_password():
    st.stop()

st.title("Анализ образовательных программ и достижений студентов")

# Функция для загрузки конфигурации
def load_config():
    config_path = Path("config.yaml")
    if not config_path.exists():
        st.error("Конфигурационный файл не найден")
        return None
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)

# Функция для загрузки последних данных
def load_latest_data():
    # Находим последние файлы с данными
    files = [f for f in os.listdir('.') if f.startswith(('students_clean_', 'events_clean_', 'student_event_relations_'))]
    if not files:
        st.error("Файлы с данными не найдены. Пожалуйста, сначала запустите main.py")
        return None, None, None
    
    # Получаем временные метки из имен файлов
    timestamps = set()
    for file in files:
        if '_' in file:
            timestamp = file.split('_')[-1].replace('.xlsx', '')
            timestamps.add(timestamp)
    
    if not timestamps:
        st.error("Не удалось найти файлы с данными")
        return None, None, None
    
    # Берем последнюю временную метку
    latest_timestamp = max(timestamps)
    
    # Загружаем данные
    students_df = pd.read_excel(f'students_clean_20250530_011651.xlsx')
    events_df = pd.read_excel(f'events_clean_20250530_011651.xlsx')
    relations_df = pd.read_excel(f'student_event_relations_20250530_011651.xlsx')
    
    return students_df, events_df, relations_df

# Загружаем данные
students_df, events_df, relations_df = load_latest_data()

if students_df is not None and events_df is not None and relations_df is not None:
    # Создаем боковую панель с фильтрами
    st.sidebar.title("Фильтры")
    
    # Фильтр по типу мероприятия
    event_type = st.sidebar.selectbox(
        "Тип мероприятия",
        ["Все", "Курс", "Соревнование"]
    )
    
    # Фильтр по году
    years = sorted(events_df['Год'].unique())
    selected_years = st.sidebar.multiselect(
        "Год",
        years,
        default=years
    )
    
    # Фильтр по региону
    regions = sorted(students_df['РЕГИОН'].dropna().unique())
    selected_regions = st.sidebar.multiselect(
        "Регион",
        regions,
        default=regions
    )
    
    # Фильтр по городу
    cities = sorted(students_df['ГОРОД'].dropna().unique())
    selected_cities = st.sidebar.multiselect(
        "Город",
        cities,
        default=cities
    )
    
    # Применяем фильтры
    filtered_events = events_df[
        (event_type == "Все" or events_df['Тип мероприятия'] == event_type) &
        (events_df['Год'].isin(selected_years))
    ]
    
    filtered_students = students_df[
        (students_df['РЕГИОН'].isin(selected_regions)) &
        (students_df['ГОРОД'].isin(selected_cities))
    ]
    
    # Создаем вкладки
    tab1, tab2, tab3 = st.tabs(["Анализ мероприятий", "Трек студента", "Анализ эффективности"])
    
    with tab1:
        st.header("Анализ мероприятий")
        
        # Подсчет участников для каждого мероприятия
        event_participants = relations_df.groupby('id_мероприятия').size().reset_index(name='Количество участников')
        event_participants = event_participants.merge(filtered_events, left_on='id_мероприятия', right_on='id')
        
        # Сортировка мероприятий
        sort_by = st.selectbox(
            "Сортировка",
            ["По количеству участников (по убыванию)", "По количеству участников (по возрастанию)", "По году (по убыванию)", "По году (по возрастанию)"]
        )
        
        if "по убыванию" in sort_by:
            ascending = False
        else:
            ascending = True
            
        if "количеству участников" in sort_by:
            event_participants = event_participants.sort_values('Количество участников', ascending=ascending)
        else:
            event_participants = event_participants.sort_values('Год', ascending=ascending)
        
        # Отображение таблицы мероприятий
        st.dataframe(
            event_participants[['Мероприятие', 'Тип мероприятия', 'Год', 'Количество участников']],
            use_container_width=True
        )
        
        # График количества участников по мероприятиям
        fig = px.bar(
            event_participants,
            x='Мероприятие',
            y='Количество участников',
            color='Тип мероприятия',
            title='Количество участников по мероприятиям'
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Анализ повторяющихся мероприятий
        st.subheader("Повторяющиеся мероприятия")
        recurring_events = event_participants.groupby('Мероприятие').size().reset_index(name='Количество проведений')
        recurring_events = recurring_events[recurring_events['Количество проведений'] > 1]
        
        if not recurring_events.empty:
            st.dataframe(recurring_events, use_container_width=True)
            
            # График динамики участников для повторяющихся мероприятий
            for event in recurring_events['Мероприятие']:
                event_data = event_participants[event_participants['Мероприятие'] == event]
                # Сортируем данные по году
                event_data = event_data.sort_values('Год')
                fig = px.line(
                    event_data,
                    x='Год',
                    y='Количество участников',
                    title=f'Динамика участников: {event}'
                )
                # Добавляем настройку оси Y
                fig.update_layout(
                    yaxis=dict(
                        range=[0, event_data['Количество участников'].max() * 1.1]  # Добавляем 10% отступа сверху
                    )
                )
                st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        st.header("Трек студента")
        
        # Выбор студента
        student_search = st.text_input("Поиск студента (ФИО или телефон)")
        
        if student_search:
            # Преобразуем все значения в строки и экранируем специальные символы
            student_search = re.escape(student_search)
            
            # Поиск по ФИО или телефону
            filtered_students = students_df[
                students_df['ФИО'].astype(str).str.contains(student_search, case=False, na=False) |
                students_df['ТЕЛЕФОН'].astype(str).str.contains(student_search, case=False, na=False)
            ]
        
        if not filtered_students.empty:
            selected_student = st.selectbox(
                "Выберите студента",
                filtered_students['ФИО'].tolist()
            )
            
            if selected_student:
                student_id = filtered_students[filtered_students['ФИО'] == selected_student]['id'].iloc[0]
                
                # Получаем мероприятия студента
                student_events = relations_df[relations_df['id_студента'] == student_id]
                student_events = student_events.merge(events_df, left_on='id_мероприятия', right_on='id')
                
                # Отображаем информацию о студенте
                student_info = filtered_students[filtered_students['id'] == student_id].iloc[0]
                st.subheader("Информация о студенте")
                st.write(f"ФИО: {student_info['ФИО']}")
                st.write(f"Регион: {student_info['РЕГИОН']}")
                st.write(f"Город: {student_info['ГОРОД']}")
                st.write(f"Школа: {student_info['ШКОЛА']}")
                
                # Отображаем мероприятия студента
                st.subheader("Мероприятия студента")
                st.dataframe(
                    student_events[['Мероприятие', 'Тип мероприятия', 'Год']],
                    use_container_width=True
                )
                
                # Визуализация трека студента
                fig = go.Figure()
                
                # Добавляем курсы
                courses = student_events[student_events['Тип мероприятия'] == 'Курс']
                fig.add_trace(go.Scatter(
                    x=courses['Год'],
                    y=[1] * len(courses),
                    mode='markers+text',
                    name='Курсы',
                    text=courses['Мероприятие'],
                    textposition="top center",
                    marker=dict(size=10, symbol='circle')
                ))
                
                # Добавляем соревнования
                competitions = student_events[student_events['Тип мероприятия'] == 'Соревнование']
                fig.add_trace(go.Scatter(
                    x=competitions['Год'],
                    y=[2] * len(competitions),
                    mode='markers+text',
                    name='Соревнования',
                    text=competitions['Мероприятие'],
                    textposition="top center",
                    marker=dict(size=10, symbol='star')
                ))
                
                fig.update_layout(
                    title='Трек студента',
                    yaxis=dict(
                        showticklabels=False,
                        range=[0, 3]
                    ),
                    showlegend=True
                )
                
                st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.header("Анализ эффективности")
        
        # Анализ связи между курсами и соревнованиями
        st.subheader("Связь между курсами и соревнованиями")
        
        # Получаем все пары курс-соревнование для каждого студента
        course_competition_pairs = []
        
        for student_id in relations_df['id_студента'].unique():
            student_events = relations_df[relations_df['id_студента'] == student_id]
            student_events = student_events.merge(events_df, left_on='id_мероприятия', right_on='id')
            
            # Получаем курсы и соревнования студента
            courses = student_events[student_events['Тип мероприятия'] == 'Курс']
            competitions = student_events[student_events['Тип мероприятия'] == 'Соревнование']
            
            # Анализируем каждую пару курс-соревнование
            for _, course in courses.iterrows():
                for _, competition in competitions.iterrows():
                    if course['Год'] <= competition['Год']:  # Курс должен быть до соревнования
                        course_competition_pairs.append({
                            'Студент': students_df[students_df['id'] == student_id]['ФИО'].iloc[0],
                            'Курс': course['Мероприятие'],
                            'Год курса': course['Год'],
                            'Соревнование': competition['Мероприятие'],
                            'Год соревнования': competition['Год'],
                            'Место': competition['Место']
                        })
        
        if course_competition_pairs:
            pairs_df = pd.DataFrame(course_competition_pairs)
            
            # Анализ популярных пар курс-соревнование
            popular_pairs = pairs_df.groupby(['Курс', 'Соревнование']).agg({
                'Студент': 'count',
                'Место': lambda x: {
                    'Победители': sum(1 for m in x if pd.notna(m) and m == 'Победитель'),
                    'Призеры': sum(1 for m in x if pd.notna(m) and m == 'Призер'),
                    'Не заняли места': sum(1 for m in x if pd.notna(m) and m == 'Участник')
                }
            }).reset_index()
            
            popular_pairs.columns = ['Курс', 'Соревнование', 'Количество студентов', 'Статистика мест']
            
            # Добавляем столбцы с количеством победителей, призеров и участников
            popular_pairs['Победители'] = popular_pairs['Статистика мест'].apply(lambda x: x['Победители'])
            popular_pairs['Призеры'] = popular_pairs['Статистика мест'].apply(lambda x: x['Призеры'])
            popular_pairs['Не заняли места'] = popular_pairs['Статистика мест'].apply(lambda x: x['Не заняли места'])
            
            # Сортируем по количеству победителей и призеров
            popular_pairs = popular_pairs.sort_values(['Победители', 'Призеры'], ascending=False)
            
            st.write("Популярные пары курс-соревнование:")
            st.dataframe(popular_pairs[['Курс', 'Соревнование', 'Количество студентов', 'Победители', 'Призеры', 'Не заняли места']], 
                        use_container_width=True)
            
            # Визуализация связей с учетом мест
            fig = px.scatter(
                pairs_df,
                x='Год курса',
                y='Год соревнования',
                color='Курс',
                hover_data=['Студент', 'Соревнование', 'Место'],
                title='Связь между курсами и соревнованиями по годам'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Анализ эффективности курсов
            st.subheader("Эффективность курсов")
            
            # Группируем данные по курсам
            course_effectiveness = pairs_df.groupby('Курс').agg({
                'Студент': 'count',
                'Место': lambda x: {
                    'Победители': sum(1 for m in x if pd.notna(m) and m == 'Победитель'),
                    'Призеры': sum(1 for m in x if pd.notna(m) and m == 'Призер'),
                    'Не заняли места': sum(1 for m in x if pd.notna(m) and m == 'Участник')
                }
            }).reset_index()
            
            course_effectiveness.columns = ['Курс', 'Количество студентов', 'Статистика мест']
            
            # Добавляем столбцы с количеством победителей, призеров и участников
            course_effectiveness['Победители'] = course_effectiveness['Статистика мест'].apply(lambda x: x['Победители'])
            course_effectiveness['Призеры'] = course_effectiveness['Статистика мест'].apply(lambda x: x['Призеры'])
            course_effectiveness['Не заняли места'] = course_effectiveness['Статистика мест'].apply(lambda x: x['Не заняли места'])
            
            # Добавляем процентные показатели
            course_effectiveness['% Победителей'] = (course_effectiveness['Победители'] / course_effectiveness['Количество студентов'] * 100).round(1)
            course_effectiveness['% Призеров'] = (course_effectiveness['Призеры'] / course_effectiveness['Количество студентов'] * 100).round(1)
            course_effectiveness['% Не заняли места'] = (course_effectiveness['Не заняли места'] / course_effectiveness['Количество студентов'] * 100).round(1)
            
            # Сортируем по проценту победителей и призеров
            course_effectiveness = course_effectiveness.sort_values(['% Победителей', '% Призеров'], ascending=False)
            
            st.write("Эффективность курсов (по количеству и проценту победителей и призеров):")
            st.dataframe(course_effectiveness[[
                'Курс', 'Количество студентов', 
                'Победители', '% Победителей',
                'Призеры', '% Призеров',
                'Не заняли места', '% Не заняли места'
            ]], use_container_width=True)
            
            # Визуализация эффективности курсов
            fig = px.bar(
                course_effectiveness,
                x='Курс',
                y=['% Победителей', '% Призеров', '% Не заняли места'],
                title='Эффективность курсов (в процентах)',
                barmode='group',
                labels={'value': 'Процент студентов', 'variable': 'Категория'}
            )
            st.plotly_chart(fig, use_container_width=True)
            
        else:
            st.write("Не найдено связей между курсами и соревнованиями")
