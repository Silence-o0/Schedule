from math import floor, ceil
from random import choice, randint

from data_workers import *
import copy


def fitness_soft(schedule, groups, teachers, auditoriums, week_quantity, output=False):
    score = 0

    window_group_penalty = 0
    for group in groups:
        for day in Day:
            day_lessons = sorted(
                [lesson for lesson in schedule.lessons if
                 lesson.group == group.name and lesson.day == day],
                key=lambda x: x.lesson_num
            )
            for i in range(len(day_lessons) - 1):
                if (day_lessons[i + 1].lesson_num - day_lessons[i].lesson_num) > 1:
                    window_group_penalty += 1
    score -= window_group_penalty

    window_teacher_penalty = 0
    for teacher in teachers:
        for day in Day:
            day_lessons = sorted(
                [lesson for lesson in schedule.lessons if
                 lesson.teacher == teacher.name and lesson.day == day],
                key=lambda x: x.lesson_num
            )
            for i in range(len(day_lessons) - 1):
                if (day_lessons[i + 1].lesson_num - day_lessons[i].lesson_num) > 1:
                    window_teacher_penalty += 1
    score -= window_teacher_penalty

    capacity_penalty = 0
    lessons_by_auditorium_time = defaultdict(lambda: defaultdict(list))
    for lesson in schedule.lessons:
        lessons_by_auditorium_time[lesson.auditorium][(lesson.day, lesson.lesson_num)].append(lesson)

    for auditorium_num, lessons_by_day in lessons_by_auditorium_time.items():
        for (day, lesson_num), lessons in lessons_by_day.items():
            total_students = sum(next(g.students_count for g in groups if g.name == lesson.group) for lesson in lessons)
            auditorium = next(a for a in auditoriums if a.number == auditorium_num)

            if total_students > auditorium.capacity:
                capacity_penalty = total_students - auditorium.capacity
    score -= capacity_penalty

    weekly_hours_penalty = 0
    for teacher in teachers:
        teacher_lessons = sorted(
            [lesson for lesson in schedule.lessons if lesson.teacher == teacher.name],
            key=lambda x: (x.day, x.lesson_num)
        )
        unique_slots = set()
        for lesson in teacher_lessons:
            unique_slots.add((lesson.day, lesson.lesson_num))
        total_hours = len(unique_slots) * 1.5
        if total_hours > teacher.hours:
            weekly_hours_penalty += total_hours - teacher.hours
    score -= weekly_hours_penalty

    overlearning_time_penalty = 0
    for group in groups:
        subject_dict = {}
        for lesson in schedule.lessons:
            if lesson.group == group.name:
                name = str(lesson.subject + "-" + lesson.lesson_type)
                if not name in subject_dict:
                    subject_dict[name] = 1
                else:
                    subject_dict[name] += 1
        for subject_item in subject_dict:
            split_array = subject_item.split("-")
            subj_name = split_array[0]
            subj_type = split_array[1]
            result_detail = next(
                (detail for subject in group.subjects if subject.name == subj_name
                 for detail in subject.details if detail.type == subj_type),
                None
            )
            subgroups_count = result_detail.subgroups
            if not subgroups_count:
                subgroups_count = 1
            subject_time = abs(
                (subject_dict[subject_item] * 1.5 * week_quantity) / subgroups_count - result_detail.hours)
            overlearning_time_penalty += subject_time
    overlearning_time_penalty /= week_quantity
    score -= overlearning_time_penalty

    if output:
        print("Groups windows:", -window_group_penalty)
        print("Teachers windows:", -window_teacher_penalty)
        print("Auditoriums capacity:", -capacity_penalty)
        print("Teachers hours constraints:", -weekly_hours_penalty)
        print("Groups hours overlearning/underlearning:", -overlearning_time_penalty)
        print("Final score: ", score)
        print()
    return score


def crossover(schedule1, schedule2, groups, teachers, week_quantity, auditoriums, max_lessons_per_day=4):
    new_schedule = Schedule()

    for group in groups:
        assigned_teachers = {}

        for day in Day:
            for lesson_num in range(1, max_lessons_per_day + 1):
                parent_schedule = schedule1 if random.random() < 0.5 else schedule2
                lesson = next(
                    (l for l in parent_schedule.lessons
                     if l.group == group.name and l.day == day and l.lesson_num == lesson_num),
                    None
                )

                if lesson:
                    subject = lesson.subject
                    lesson_type = lesson.lesson_type

                    if (subject, lesson_type) not in assigned_teachers:
                        assigned_teachers[(subject, lesson_type)] = lesson.teacher

                    if lesson.teacher == assigned_teachers[(subject, lesson_type)]:
                        if new_schedule.check_hard_constraints(lesson):
                            new_schedule.lessons.append(lesson)
                        else:
                            alt_schedule = schedule2 if parent_schedule == schedule1 else schedule1
                            alt_lesson = next(
                                (l for l in alt_schedule.lessons
                                 if l.group == group.name and l.day == day and l.lesson_num == lesson_num),
                                None
                            )
                            if (alt_lesson and
                                    alt_lesson.teacher == assigned_teachers[(subject, lesson_type)] and
                                    new_schedule.check_hard_constraints(alt_lesson)):
                                new_schedule.lessons.append(alt_lesson)

    mutated_schedule = mutation_fixed_group_subjects(new_schedule, groups, teachers, auditoriums, week_quantity)
    smoothed_schedule = smoothing(mutated_schedule, teachers)
    mutated_schedule = mutation_fixed_group_subjects(smoothed_schedule, groups, teachers, auditoriums, week_quantity)
    mutated_schedule = mutate_auditoriums_by_size(mutated_schedule, auditoriums, groups)
    return mutated_schedule


def mutation_fixed_group_subjects(schedule, groups, teachers, auditoriums, week_quantity):
    for group in groups:
        for subject in group.subjects:
            for detail in subject.details:
                subgroups_count = 1
                if detail.subgroups is not None:
                    subgroups_count = detail.subgroups

                for subgroup in range(1, subgroups_count + 1):
                    subgroup_name = str(str(subgroup) + "/" + str(subgroups_count))
                    subject_hours = sum(
                        1.5 for lesson in schedule.lessons
                        if lesson.group == group.name and lesson.subject == subject.name and
                        detail.type == lesson.lesson_type and (detail.subgroups == 1 or detail.subgroups is None or
                                                               lesson.subgroup == subgroup_name)
                    )
                    max_hours = detail.hours / week_quantity

                    if subject_hours > max_hours:
                        excess_hours = subject_hours - max_hours
                        lessons_to_remove = [
                            lesson for lesson in schedule.lessons
                            if lesson.group == group.name and lesson.subject == subject.name and
                               detail.type == lesson.lesson_type and (
                                       detail.subgroups == 1 or detail.subgroups is None or
                                       lesson.subgroup == subgroup_name)
                        ]
                        for _ in range(int(excess_hours / 1.5)):
                            lesson_to_random = []
                            for lesson in lessons_to_remove:
                                if lesson.lesson_num == 1 or lesson.lesson_num == 4:
                                    lesson_to_random.append(lesson)
                            if len(lesson_to_random) < 1:
                                lesson_to_random = lessons_to_remove
                            to_delete = random.choice(lesson_to_random)
                            lessons_to_remove.remove(to_delete)
                            schedule.lessons.remove(to_delete)

                    if subject_hours < max_hours / 2:
                        lack_hours = max_hours - subject_hours
                        existing_lesson = [
                            lesson for lesson in schedule.lessons
                            if lesson.group == group.name and lesson.subject == subject.name and
                               detail.type == lesson.lesson_type and (
                                       detail.subgroups == 1 or detail.subgroups is None or
                                       lesson.subgroup == subgroup_name)
                        ]
                        available_teachers = []
                        if not existing_lesson:
                            available_teachers = [
                                t
                                for t in teachers
                                if any(
                                    ts.name == subject.name and max_hours <= t.hours
                                    and detail.type in [tsd.type for tsd in ts.details]
                                    for ts in t.subjects
                                )
                            ]
                        else:
                            teacher = [t for t in teachers if t.name == existing_lesson[0].teacher]
                            available_teachers.append(teacher[0])

                        for _ in range(int(lack_hours / 1.5)):
                            days_list = list(Day)
                            random.shuffle(days_list)
                            assigned = False
                            for day in days_list:
                                for i in range(2, 4):
                                    flag = 0
                                    for lesson in schedule.lessons:
                                        if (lesson.group == group.name and lesson.day == day and
                                                lesson.lesson_num == i and (
                                                        detail.subgroups == 1 or detail.subgroups is None or
                                                        lesson.subgroup == subgroup_name)):
                                            break
                                        if (lesson.group == group.name and lesson.day == day and
                                                lesson.lesson_num == i - 1):
                                            if subgroup is not None and subgroup != 1:
                                                if lesson.subgroup == subgroup_name:
                                                    flag += 1
                                            else:
                                                flag += 1
                                        if (lesson.group == group.name and lesson.day == day and
                                                lesson.lesson_num == i + 1):
                                            if subgroup is not None and subgroup != 1:
                                                if lesson.subgroup == subgroup_name:
                                                    flag += 1
                                            else:
                                                flag += 1
                                    if flag > 1:
                                        for teacher in available_teachers:
                                            auditorium = random.choice(
                                                [
                                                    a
                                                    for a in auditoriums
                                                    if a.capacity >= group.students_count
                                                ]
                                            )
                                            if not auditorium:
                                                auditorium = random.choice(
                                                    [
                                                        a
                                                        for a in auditoriums
                                                    ]
                                                )

                                            if detail.type == "Lec" or subgroups_count == 1:
                                                subgroup_name = None

                                            lesson = Lesson(
                                                day,
                                                i,
                                                teacher,
                                                detail.type,
                                                subject,
                                                group,
                                                auditorium,
                                                subgroup_name,
                                            )

                                            if detail.type == "Lec":
                                                result_flag, new_lesson = schedule.set_shared_lec(lesson, teacher,
                                                                                                  subject, group,
                                                                                                  auditoriums)
                                                if result_flag:
                                                    lesson = new_lesson

                                            if schedule.check_hard_constraints(lesson):
                                                schedule.lessons.append(lesson)
                                                assigned = True
                                                break
                            loop_num = 0
                            while not assigned and loop_num < 100:
                                teacher = random.choice(available_teachers)
                                day = random.choice(days_list)
                                lesson_num = random.randint(1, 4)
                                auditorium = random.choice(
                                    [
                                        a
                                        for a in auditoriums
                                        if a.capacity >= group.students_count
                                    ]
                                )
                                if not auditorium:
                                    auditorium = random.choice(
                                        [
                                            a
                                            for a in auditoriums
                                        ]
                                    )

                                if detail.type == "Lec" or subgroups_count == 1:
                                    subgroup_name = None

                                lesson = Lesson(
                                    day,
                                    lesson_num,
                                    teacher,
                                    detail.type,
                                    subject,
                                    group,
                                    auditorium,
                                    subgroup_name,
                                )

                                if detail.type == "Lec":
                                    result_flag, new_lesson = schedule.set_shared_lec(lesson, teacher, subject,
                                                                                      group, auditoriums)
                                    if result_flag:
                                        lesson = new_lesson

                                loop_num += 1
                                if schedule.check_hard_constraints(lesson):
                                    schedule.lessons.append(lesson)
                                    assigned = True
    return schedule


def smoothing(new_schedule, teachers):
    for teacher in teachers:
        teacher_lessons = sorted(
            [lesson for lesson in new_schedule.lessons if lesson.teacher == teacher.name],
            key=lambda x: (x.day, x.lesson_num)
        )

        unique_slots = set()
        for lesson in teacher_lessons:
            unique_slots.add((lesson.day, lesson.lesson_num))

        teacher_hours = len(unique_slots) * 1.5

        if teacher_hours > teacher.hours:
            excess_hours = teacher_hours - teacher.hours

            lessons_by_subject_type = {}
            for lesson in new_schedule.lessons:
                if lesson.teacher == teacher.name:
                    key = (lesson.subject, lesson.lesson_type, lesson.group)
                    if key not in lessons_by_subject_type:
                        lessons_by_subject_type[key] = []
                    lessons_by_subject_type[key].append(lesson)

            while excess_hours > 0 and lessons_by_subject_type:
                random_key = random.choice(list(lessons_by_subject_type.keys()))
                random_group_lessons = lessons_by_subject_type[random_key]

                if excess_hours >= 1.5 * (len(random_group_lessons) - 1):
                    for lesson in random_group_lessons:
                        new_schedule.lessons.remove(lesson)
                    excess_hours -= 1.5 * len(random_group_lessons)

                del lessons_by_subject_type[random_key]

            if excess_hours <= 0:
                break

    return new_schedule


def mutate_auditoriums_by_size(schedule, auditoriums, groups):
    new_schedule = copy.deepcopy(schedule)
    lessons = new_schedule.lessons

    for _ in range(5):
        for _ in range(20):
            lesson1, lesson2 = random.sample(lessons, 2)

            if (
                    lesson1.day == lesson2.day and
                    lesson1.lesson_num == lesson2.lesson_num and
                    lesson1.auditorium != lesson2.auditorium
            ):
                break
        else:
            return new_schedule

        total_students_lesson1 = sum(
            group.students_count for group in groups if group.name == lesson1.group
        )
        total_students_lesson2 = sum(
            group.students_count for group in groups if group.name == lesson2.group
        )

        auditorium1 = next((a for a in auditoriums if a.number == lesson1.auditorium), None)
        auditorium2 = next((a for a in auditoriums if a.number == lesson2.auditorium), None)

        if auditorium1 and auditorium2:
            if (
                    (total_students_lesson1 > total_students_lesson2 and auditorium1.capacity < auditorium2.capacity) or
                    (total_students_lesson1 < total_students_lesson2 and auditorium1.capacity > auditorium2.capacity)
            ):
                lesson1.auditorium, lesson2.auditorium = lesson2.auditorium, lesson1.auditorium
    return new_schedule


def calculate_similarity(schedule1, schedule2):
    lessons1 = set(
        (l.day, l.lesson_num, l.teacher, l.group, l.lesson_type, l.subject, l.auditorium)
        for l in schedule1.lessons
    )
    lessons2 = set(
        (l.day, l.lesson_num, l.teacher, l.group, l.lesson_type, l.subject, l.auditorium)
        for l in schedule2.lessons
    )

    intersection = len(lessons1 & lessons2)
    union = len(lessons1 | lessons2) // 2
    return intersection / union if union > 0 else 0.0


def group_schedules(schedule_pool, similarity_threshold=1.0):
    num_schedules = len(schedule_pool)
    visited = set()
    clusters = []

    for i in range(num_schedules):
        if i not in visited:
            cluster = [i]
            visited.add(i)
            for j in range(num_schedules):
                if j != i and j not in visited:
                    similarity = calculate_similarity(schedule_pool[i][0], schedule_pool[j][0])
                    if similarity >= similarity_threshold:
                        cluster.append(j)
                        visited.add(j)
            clusters.append(cluster)
    return clusters


def predator_approach(schedules_tuple, clusters, retain_count=8):
    new_schedules_tuples = []

    for cluster in clusters:
        if len(cluster) > retain_count:
            sorted_cluster = sorted(
                cluster, key=lambda idx: schedules_tuple[idx][1], reverse=True
            )
            best_indices = sorted_cluster[:retain_count]
        else:
            best_indices = cluster

        for idx in best_indices:
            new_schedules_tuples.append(schedules_tuple[idx])

    return new_schedules_tuples


def rain_effect(schedules_collection, max_num, teachers, auditoriums, groups, change_num=10, max_attempts=50):
    clusters = group_schedules(schedules_collection, similarity_threshold=0.8)

    sorted_clusters = sorted(clusters, key=lambda x: len(x))
    num_clusters_to_rain = len(sorted_clusters) // 2
    target_clusters = sorted_clusters[:num_clusters_to_rain]

    new_schedules = copy.deepcopy(schedules_collection)

    for cluster in target_clusters:
        base_schedules_idx = random.choices(cluster, k=ceil(len(cluster) / 2))
        for base_schedule_idx in base_schedules_idx:
            while len(new_schedules) <= max_num:
                new_schedule = copy.deepcopy(schedules_collection[base_schedule_idx][0])
                changes_made = 0
                attempts = 0

                while attempts < max_attempts:
                    action = random.choice([0, 1])

                    if action == 0:
                        if change_teacher(new_schedule, teachers):
                            changes_made += 1
                    else:
                        if change_auditorium(new_schedule, auditoriums, groups):
                            changes_made += 1

                    attempts += 1
                if changes_made >= change_num:
                    new_schedules.append(new_schedule)

    return new_schedules


def change_teacher(schedule, teachers):
    lesson = random.choice(schedule.lessons)
    copy_lesson = copy.deepcopy(lesson)
    eligible_teachers = [
        teacher for teacher in teachers
        if lesson.subject in teacher.subjects and teacher.name != lesson.teacher
    ]

    if eligible_teachers:
        schedule.lessons.remove(lesson)
        new_teacher = random.choice(eligible_teachers)
        copy_lesson.teacher = new_teacher.name
        if schedule.check_hard_constraints(copy_lesson):
            schedule.lessons.append(copy_lesson)
            return True
        schedule.lessons.append(lesson)
    return False


def change_auditorium(schedule, auditoriums, groups):
    lesson = random.choice(schedule.lessons)
    copy_lesson = copy.deepcopy(lesson)
    eligible_auditoriums = [
        auditorium for auditorium in auditoriums
        if auditorium.capacity >= sum(group.students_count for group in groups if group.name == lesson.group)
    ]

    if eligible_auditoriums:
        schedule.lessons.remove(lesson)
        new_auditorium = random.choice(eligible_auditoriums)
        copy_lesson.auditorium = new_auditorium.number
        if schedule.check_hard_constraints(copy_lesson):
            schedule.lessons.append(copy_lesson)
            return True
        schedule.lessons.append(lesson)
    return False


if __name__ == "__main__":
    file_path = "schedule_data.xlsx"
    init_groups, init_teachers, init_auditoriums = load_data_from_excel(file_path)
    gen_subjects, gen_teachers, gen_groups, gen_auditoriums = test_generate(
        5, 5, 10, 8
        # 0, 0, 0, 0
    )
    week_quantity = 14

    groups = gen_groups + init_groups
    teachers = gen_teachers + init_teachers
    auditoriums = gen_auditoriums + init_auditoriums

    schedules_collection = []

    specimen_num = 100
    iter_num = 400

    try:
        for i in range(specimen_num):
            schedule = Schedule()
            schedule.generate_schedule(
                groups=groups, teachers=teachers, auditoriums=auditoriums, week_quantity=week_quantity
            )
            fitness_score = fitness_soft(schedule, groups, teachers, auditoriums, week_quantity, False)
            schedules_collection.append((schedule, fitness_score))
    except Exception as e:
        print(f"{e}")
        exit(0)

    num_iter_no_change = (None, 0)
    for i in range(iter_num):

        clusters = group_schedules(schedules_collection, similarity_threshold=0.8)
        # for idx, cluster in enumerate(clusters):
        #     if len(cluster) > 1:
        #         print(f"Cluster {idx}: {cluster}")

        clean_schedule_collection = predator_approach(schedules_collection, clusters,
                                                      retain_count=max(floor(0.1 * specimen_num), 4))

        sorted_items = sorted(
            clean_schedule_collection,
            key=lambda x: x[1],
            reverse=True
        )

        half_size = specimen_num // 2
        if half_size < len(sorted_items):
            sorted_items = sorted_items[:half_size]
        schedules_collection = sorted_items


        half_num = len(sorted_items) // 4
        if (sorted_items[-half_num][1] - sorted_items[half_num][1]) < 2 and (
                num_iter_no_change[0] == sorted_items[-half_num][1]
                or num_iter_no_change[0] is None):
            num_iter_no_change = (sorted_items[-half_num][1], num_iter_no_change[1] + 1)
            print("The results change weakly")
        else:
            num_iter_no_change = (sorted_items[-half_num][1], 0)


        if num_iter_no_change[1] == 5:
            rain_effect(schedules_collection, specimen_num, teachers, auditoriums, groups, change_num=10, max_attempts=200)
            num_iter_no_change = (None, 0)

        print()
        print(f"Best of iter {i}: ")
        for fi in range(0, len(sorted_items)):
            print(sorted_items[fi][1])

        for j in range(specimen_num - len(sorted_items)):
            random_parent1 = random.randint(0, half_size - 1)
            random_parent2 = random.randint(0, half_size - 1)
            while random_parent1 == random_parent2:
                random_parent2 = random.randint(0, half_size - 1)

            child_schedule = crossover(sorted_items[random_parent1][0],
                                       sorted_items[random_parent2][0], groups, teachers, week_quantity, auditoriums)
            fitness_score = fitness_soft(child_schedule, groups, teachers, auditoriums, week_quantity, False)
            schedules_collection.append((child_schedule, fitness_score))

    schedules_collection = sorted(
        schedules_collection,
        key=lambda x: x[1],
        reverse=True
    )

    print()
    print("BEST:")
    print("Hard constraints:", schedules_collection[0][0].hard_constraints_schedule_check())
    fitness_soft(schedules_collection[0][0], groups, teachers, auditoriums, week_quantity, True)
    export_schedule_to_excel(schedules_collection[0][0], groups, auditoriums, "schedule.xlsx")
