import os
import sys
import time
import traceback

def setup_logging():
    """设置日志记录"""
    try:
        # 将日志目录改到用户目录下
        log_dir = os.path.expanduser('~/Library/Logs/Word2PDF')
        if not os.path.exists(log_dir):
            try:
                os.makedirs(log_dir, exist_ok=True)
            except Exception:
                return None
        
        log_file = os.path.join(log_dir, f"word2pdf_{time.strftime('%Y%m%d')}.log")
        return log_file
    except Exception:
        return None

def log_system_info():
    """记录系统信息"""
    import platform
    
    log_info = f"=== Word2PDF 系统信息 === 时间: {time.strftime('%Y-%m-%d %H:%M:%S')} 系统: {platform.system()} {platform.release()}"
    print(log_info)

def log_exception(exc_type, exc_value, exc_traceback):
    """记录未捕获的异常"""
    error_info = f"错误: {exc_type.__name__} - {exc_value}"
    print(error_info)
    
    log_file = setup_logging()
    if log_file:
        try:
            with open(log_file, 'a') as f:
                f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] {error_info}\n")
                f.write(''.join(traceback.format_tb(exc_traceback)))
        except:
            pass