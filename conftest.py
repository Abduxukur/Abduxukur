import config
from config import LOG_FILE_DIR
from config.VarConfig import projectPath
from util.read_excel import MyExcel
import datetime
import json
import os
import time
import pytest
from jinja2 import Environment, FileSystemLoader

test_result = {
    "title": "",
    "tester": "",
    "desc": "",
    "cases": {},
    'rerun': 0,
    "failed": 0,
    "passed": 0,
    "skipped": 0,
    "error": 0,
    "start_time": 0,
    "run_time": 0,
    "begin_time": "",
    "all": 0,
    "testModules": set()
}


def pytest_make_parametrize_id(config, val, argname):
    if isinstance(val, dict):
        return val.get('title') or val.get('desc')


def pytest_runtest_logreport(report):
    report.duration = '{:.6f}'.format(report.duration)
    test_result['testModules'].add(report.fileName)
    # if True:
    if report.when == 'call':
        test_result[report.outcome] += 1
        test_result["cases"][report.nodeid] = report
    elif report.outcome == 'failed':
        report.outcome = 'error'
        test_result['error'] += 1
        test_result["cases"][report.nodeid] = report
    elif report.outcome == 'skipped':
        test_result[report.outcome] += 1
        test_result["cases"][report.nodeid] = report


def pytest_sessionstart(session):
    start_ts = datetime.datetime.now()
    test_result["start_time"] = start_ts.timestamp()
    test_result["begin_time"] = start_ts.strftime("%Y-%m-%d %H:%M:%S")


def handle_history_data(report_dir, test_result):
    """
    处理历史数据
    :return:
    """
    try:
        with open(os.path.join(report_dir, 'history.json'), 'r', encoding='utf-8') as f:
            history = json.load(f)
    except:
        history = []
    history.append({'success': test_result['passed'],
                    'all': test_result['all'],
                    'fail': test_result['failed'],
                    'skip': test_result['skipped'],
                    'error': test_result['error'],
                    'runtime': test_result['run_time'],
                    'begin_time': test_result['begin_time'],
                    'pass_rate': test_result['pass_rate'],
                    })

    with open(os.path.join(report_dir, 'history.json'), 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=True)
    return history


def pytest_sessionfinish(session):
    """在整个测试运行完成之后调用的钩子函数,可以在此处生成测试报告"""
    report = session.config.getoption('--Report')
    title = session.config.getoption('--Title')
    tester = session.config.getoption('--Owner')
    desc = session.config.getoption('--Desc')
    template = session.config.getoption('--Template')

    if report:
        test_result['title'] = title
        test_result['tester'] = tester
        test_result['desc'] = desc
        templates_name = template
        file_name = 'pytest_report.html'
    else:
        return
    if os.path.isdir('reports'):
        pass
    else:
        os.mkdir('reports')
    file_name = os.path.join(LOG_FILE_DIR, file_name)
    test_result["run_time"] = '{:.3f} S'.format(time.time() - test_result["start_time"])
    test_result['all'] = len(test_result['cases'])
    if test_result['all'] != 0:
        test_result['pass_rate'] = '{:.3f}'.format(test_result['passed'] / test_result['all'] * 100)
    else:
        test_result['pass_rate'] = 0
    # 保存历史数据
    test_result['history'] = handle_history_data(config.REPORT_PATH, test_result)
    # 渲染报告
    template_path = os.path.join(projectPath, 'pytestTestreport/templates')
    env = Environment(loader=FileSystemLoader(template_path))

    if templates_name == '2':
        template = env.get_template('templates2.html')
    else:
        template = env.get_template('templates.html')
    report = template.render(test_result)
    with open(file_name, 'wb') as f:
        f.write(report.encode('utf-8'))


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    fixture_extras = getattr(item.config, "extras", [])
    plugin_extras = getattr(report, "extra", [])
    report.extra = fixture_extras + plugin_extras
    report.fileName = item.location[0]
    if hasattr(item, 'callspec'):
        report.desc = item.callspec.params['case']['desc']
        report.method = item.callspec.params['case']['story']
    else:
        report.desc = item._obj.__doc__
        report.method = item.location[2].split('[')[0]


def pytest_addoption(parser):
    parser.addoption(
        "--Report",
        action="store",
        default=True,
        help="assign which file to use",
    )
    parser.addoption(
        "--Owner",
        action="store",
        default='Jumai',
        help="assign which file to use",
    )
    parser.addoption(
        "--Title",
        action="store",
        default='Debug Case',
        help="assign which file to use",
    )
    parser.addoption(
        "--Template",
        action="store",
        default=None,
        help="assign which file to use",
    )
    parser.addoption(
        "--Desc",
        action="store",
        default='小小的测试report',
        help="assign which file to use",
    )
    parser.addoption(
        "--case_tag",
        action="store",
        default=None,
        help="assign which env to use",
    )
    parser.addoption(
        "--excel",
        action="store",
        default='sample.xlsx',
        help="assign which file to use",
    )


@pytest.hookimpl(tryfirst=True)
def pytest_generate_tests(metafunc):
    tag = metafunc.config.getoption('--case_tag')
    file = metafunc.config.getoption('--excel')
    cases = MyExcel(file).read_data(case_tag=tag)
    ilist = []
    title = []
    for ts in cases:
        ilist.append(ts['story'])
        title.append(ts['story'])
    if "case" in metafunc.fixturenames:
        metafunc.parametrize("case", cases, indirect=True, ids=ilist)


@pytest.fixture
def case(request):
    if request.param is not None:
        return request.param
    else:
        raise ValueError("invalid internal test config")


@pytest.fixture(scope="session")
def case_tag(pytestconfig):
    return pytestconfig.getoption('--case_tag')


@pytest.fixture(scope="session")
def template(request):
    return request.config.group.getoption('--Template')


@pytest.fixture(scope="session")
def desc(request):
    return request.config.group.getoption('--Desc')


@pytest.fixture(scope="session")
def tester(request):
    return request.config.group.getoption('--owner')


@pytest.fixture(scope="session")
def title(request):
    return request.config.group.getoption('--Title')


@pytest.fixture(scope="session")
def excel(pytestconfig):
    return pytestconfig.getoption('--excel')
